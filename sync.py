"""
Calendar Sync - synchronizes events between Office 365 and iCloud calendars.
Runs twice daily via Windows Task Scheduler.
"""

import os
import sys
import uuid
import logging
from datetime import datetime, timedelta, timezone

import html
from zoneinfo import ZoneInfo
import yaml
import msal
import requests
import caldav
from dotenv import load_dotenv

VIENNA = ZoneInfo("Europe/Vienna")

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("sync.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

SYNC_WINDOW_DAYS = 30
SYNC_TAG = "[synced]"

# ---------------------------------------------------------------------------
# Office 365 / Microsoft Graph
# ---------------------------------------------------------------------------

_token_cache: dict[str, str] = {}


def get_graph_token(tenant_key: str) -> str:
    """Holt einen Graph-Token für den angegebenen Tenant (layest oder invester)."""
    if tenant_key in _token_cache:
        return _token_cache[tenant_key]

    prefix = f"O365_{tenant_key.upper()}"
    client_id = os.environ[f"{prefix}_CLIENT_ID"]
    client_secret = os.environ[f"{prefix}_CLIENT_SECRET"]
    tenant_id = os.environ[f"{prefix}_TENANT_ID"]

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Graph token error [{tenant_key}]: {result.get('error_description')}")
    _token_cache[tenant_key] = result["access_token"]
    return result["access_token"]


def fetch_o365_events(account_email: str, tenant_key: str) -> list[dict]:
    token = get_graph_token(tenant_key)
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=SYNC_WINDOW_DAYS)
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    url = (
        f"https://graph.microsoft.com/v1.0/users/{account_email}/calendarView"
        f"?startDateTime={now.strftime(fmt)}&endDateTime={end.strftime(fmt)}"
        f"&$select=id,subject,start,end,body,isAllDay,showAs,categories"
        f"&$top=100"
    )
    headers = {"Authorization": f"Bearer {token}", "Prefer": 'outlook.timezone="UTC"'}
    events = []
    while url:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        events.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    log.info("Office 365 [%s]: %d Termine geladen", account_email, len(events))
    return events


def create_o365_event(account_email: str, tenant_key: str, event: dict) -> None:
    token = get_graph_token(tenant_key)
    url = f"https://graph.microsoft.com/v1.0/users/{account_email}/events"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.post(url, headers=headers, json=event, timeout=30)
    resp.raise_for_status()
    log.info("Office 365 [%s]: Termin erstellt '%s'", account_email, event.get("subject"))


def delete_o365_synced_events(account_email: str, tenant_key: str) -> None:
    token = get_graph_token(tenant_key)
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=SYNC_WINDOW_DAYS)
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    url = (
        f"https://graph.microsoft.com/v1.0/users/{account_email}/calendarView"
        f"?startDateTime={now.strftime(fmt)}&endDateTime={end.strftime(fmt)}"
        f"&$select=id,subject,body&$top=100"
    )
    headers = {"Authorization": f"Bearer {token}", "Prefer": 'outlook.timezone="UTC"'}
    while url:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        for ev in data.get("value", []):
            body = ev.get("body", {}).get("content", "")
            if SYNC_TAG in body:
                del_url = f"https://graph.microsoft.com/v1.0/users/{account_email}/events/{ev['id']}"
                requests.delete(del_url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
                log.info("Office 365 [%s]: Sync-Termin gelöscht '%s'", account_email, ev.get("subject"))
        url = data.get("@odata.nextLink")


# ---------------------------------------------------------------------------
# iCloud CalDAV
# ---------------------------------------------------------------------------

def _get_icloud_calendar(calendar_name: str) -> caldav.Calendar:
    username = os.environ["ICLOUD_USERNAME"]
    password = os.environ["ICLOUD_APP_PASSWORD"]
    client = caldav.DAVClient(
        url="https://caldav.icloud.com",
        username=username,
        password=password,
    )
    principal = client.principal()
    for cal in principal.calendars():
        if html.unescape(cal.name) == calendar_name:
            return cal
    available = [html.unescape(c.name) for c in principal.calendars()]
    raise ValueError(f"iCloud-Kalender '{calendar_name}' nicht gefunden. Verfügbar: {available}")


def fetch_icloud_events(calendar_name: str) -> list[dict]:
    cal = _get_icloud_calendar(calendar_name)
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=SYNC_WINDOW_DAYS)
    events = []
    for item in cal.date_search(start=now, end=end, expand=True):
        vevent = item.vobject_instance.vevent
        events.append({
            "uid": str(vevent.uid.value),
            "subject": str(vevent.summary.value) if hasattr(vevent, "summary") else "(kein Titel)",
            "start": vevent.dtstart.value,
            "end": vevent.dtend.value if hasattr(vevent, "dtend") else None,
            "description": str(vevent.description.value) if hasattr(vevent, "description") else "",
        })
    log.info("iCloud [%s]: %d Termine geladen", calendar_name, len(events))
    return events


def create_icloud_event(calendar_name: str, subject: str, start: datetime, end: datetime, description: str = "") -> None:
    cal = _get_icloud_calendar(calendar_name)
    fmt = "%Y%m%dT%H%M%SZ"
    ical = (
        "BEGIN:VCALENDAR\r\n"
        "VERSION:2.0\r\n"
        "BEGIN:VEVENT\r\n"
        f"UID:{uuid.uuid4()}\r\n"
        f"SUMMARY:{subject}\r\n"
        f"DTSTART:{start.strftime(fmt)}\r\n"
        f"DTEND:{end.strftime(fmt)}\r\n"
        f"DESCRIPTION:{description}\r\n"
        "END:VEVENT\r\n"
        "END:VCALENDAR\r\n"
    )
    cal.save_event(ical)
    log.info("iCloud [%s]: Termin erstellt '%s'", calendar_name, subject)


def delete_icloud_synced_events(calendar_name: str) -> None:
    cal = _get_icloud_calendar(calendar_name)
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=SYNC_WINDOW_DAYS)
    for item in cal.date_search(start=now, end=end, expand=True):
        vevent = item.vobject_instance.vevent
        desc = str(vevent.description.value) if hasattr(vevent, "description") else ""
        if SYNC_TAG in desc:
            item.delete()
            subject = str(vevent.summary.value) if hasattr(vevent, "summary") else "?"
            log.info("iCloud [%s]: Sync-Termin gelöscht '%s'", calendar_name, subject)


# ---------------------------------------------------------------------------
# Kalender laden / bereinigen (generisch über config)
# ---------------------------------------------------------------------------

def load_all_sources(cal_cfg: dict) -> dict[str, list[dict]]:
    sources = {}
    for key, cal in cal_cfg.items():
        if cal.get("disabled"):
            log.info("Kalender [%s] ist deaktiviert — übersprungen.", key)
            continue
        try:
            if cal["type"] == "o365":
                sources[key] = fetch_o365_events(cal["account"], cal["tenant_key"])
            elif cal["type"] == "icloud":
                sources[key] = fetch_icloud_events(cal["calendar_name"])
        except Exception as e:
            log.error("Fehler beim Laden von [%s]: %s", key, e)
            sources[key] = []
    return sources


def cleanup_synced_events(cal_cfg: dict) -> None:
    for key, cal in cal_cfg.items():
        if cal.get("disabled") or cal.get("readonly"):
            continue
        try:
            if cal["type"] == "o365":
                delete_o365_synced_events(cal["account"], cal["tenant_key"])
            elif cal["type"] == "icloud":
                delete_icloud_synced_events(cal["calendar_name"])
        except Exception as e:
            log.error("Fehler beim Bereinigen von [%s]: %s", key, e)


# ---------------------------------------------------------------------------
# Regel-Engine
# ---------------------------------------------------------------------------

def _normalize_dt(dt) -> datetime:
    if isinstance(dt, datetime):
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    return datetime(dt.year, dt.month, dt.day, tzinfo=timezone.utc)


def _parse_o365_dt(dt_str: str) -> datetime:
    # Graph API liefert mit Prefer:UTC keine Z-Suffix — wir hängen UTC explizit an
    clean = dt_str.split(".")[0].rstrip("Z")
    return datetime.fromisoformat(clean).replace(tzinfo=timezone.utc)


def _extract_event(ev: dict, source_type: str) -> tuple[str, datetime, datetime]:
    if source_type == "o365":
        subject = ev.get("subject", "(kein Titel)")
        start = _parse_o365_dt(ev["start"]["dateTime"])
        end = _parse_o365_dt(ev["end"]["dateTime"])
    else:
        subject = ev.get("subject", "(kein Titel)")
        start = _normalize_dt(ev["start"])
        end = _normalize_dt(ev["end"]) if ev.get("end") else start + timedelta(hours=1)
    return subject, start, end


def _apply_filters(events: list, flt: dict, source_type: str, cal_cfg: dict, source_key: str) -> list:
    if "categories" in flt:
        events = [e for e in events if any(c in e.get("categories", []) for c in flt["categories"])]

    if "subject_starts_with" in flt:
        prefix = flt["subject_starts_with"].lower()
        filtered = []
        for e in events:
            subj = e.get("subject", "") if source_type == "o365" else e.get("subject", "")
            if subj.lower().startswith(prefix):
                filtered.append(e)
        events = filtered

    if "time_outside" in flt:
        start_hour = flt["time_outside"]["start_hour"]
        end_hour = flt["time_outside"]["end_hour"]
        filtered = []
        for e in events:
            # Ganztägige Termine ignorieren
            if source_type == "o365" and e.get("isAllDay"):
                continue
            _, start, _ = _extract_event(e, source_type)
            vienna_hour = start.astimezone(VIENNA).hour
            if vienna_hour < start_hour or vienna_hour >= end_hour:
                filtered.append(e)
        events = filtered

    return events


def apply_rules(rules: list, sources: dict, cal_cfg: dict) -> None:
    if not rules:
        log.info("Keine Regeln konfiguriert — bitte config.yaml befüllen.")
        return

    for rule in rules:
        name = rule.get("name", "Unbenannte Regel")
        source_key = rule["source"]
        target_key = rule["target"]
        action = rule.get("action", "block")
        subject_override = rule.get("subject_override")
        subject_prefix = rule.get("subject_prefix", "")
        make_private = rule.get("private", False)

        target_cal = cal_cfg.get(target_key, {})
        if target_cal.get("readonly") or target_cal.get("disabled"):
            log.warning("Regel '%s' übersprungen: Ziel '%s' ist read-only/deaktiviert.", name, target_key)
            continue

        log.info("Regel: %s (%s → %s, Aktion: %s)", name, source_key, target_key, action)

        events = sources.get(source_key, [])
        source_type = cal_cfg.get(source_key, {}).get("type", "o365")

        events = _apply_filters(events, rule.get("filter", {}), source_type, cal_cfg, source_key)
        log.info("  %d Termine nach Filter", len(events))

        for ev in events:
            original_subject, start, end = _extract_event(ev, source_type)
            if action == "block":
                subject = "Besetzt"
            else:
                base = subject_override or original_subject
                subject = f"{subject_prefix}{base}"
            description = f"{SYNC_TAG} Automatisch synchronisiert von: {original_subject}"

            try:
                if target_cal["type"] == "o365":
                    payload = {
                        "subject": subject,
                        "start": {"dateTime": start.isoformat(), "timeZone": "UTC"},
                        "end": {"dateTime": end.isoformat(), "timeZone": "UTC"},
                        "body": {"contentType": "text", "content": description},
                        "showAs": "busy",
                        "sensitivity": "private" if make_private else "normal",
                    }
                    create_o365_event(target_cal["account"], target_cal["tenant_key"], payload)
                else:
                    create_icloud_event(target_cal["calendar_name"], subject, start, end, description)
            except Exception as e:
                log.error("Fehler beim Erstellen von '%s': %s", original_subject, e)


# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------

def main():
    log.info("=== Calendar Sync gestartet: %s ===", datetime.now().strftime("%Y-%m-%d %H:%M"))

    with open("config.yaml", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    cal_cfg = cfg.get("calendars", {})
    rules = cfg.get("rules", [])

    log.info("--- Alte Sync-Termine bereinigen ---")
    cleanup_synced_events(cal_cfg)

    log.info("--- Termine laden ---")
    sources = load_all_sources(cal_cfg)

    log.info("--- Regeln anwenden ---")
    apply_rules(rules, sources, cal_cfg)

    log.info("=== Sync abgeschlossen ===")


if __name__ == "__main__":
    main()
