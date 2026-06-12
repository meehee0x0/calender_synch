# Richtet Windows Task Scheduler ein: sync.py läuft um 9:30 und 16:30 Uhr täglich
$pythonPath = (Get-Command python).Source
$scriptPath = "$PSScriptRoot\sync.py"
$workDir    = $PSScriptRoot

$action  = New-ScheduledTaskAction -Execute $pythonPath -Argument $scriptPath -WorkingDirectory $workDir
$trigger1 = New-ScheduledTaskTrigger -Daily -At "09:30"
$trigger2 = New-ScheduledTaskTrigger -Daily -At "16:30"
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -StartWhenAvailable

Register-ScheduledTask `
    -TaskName "CalendarSync" `
    -Action $action `
    -Trigger $trigger1, $trigger2 `
    -Settings $settings `
    -Description "Synchronisiert Office 365 und iCloud Kalender" `
    -RunLevel Highest `
    -Force

Write-Host "Task 'CalendarSync' eingerichtet: täglich 09:30 und 16:30 Uhr." -ForegroundColor Green
