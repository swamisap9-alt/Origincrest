# Registers the three workflows in Windows Task Scheduler (folder "OriginCrest").
#   powershell -ExecutionPolicy Bypass -File scripts\install_tasks.ps1
# Re-run any time to change times; -Force replaces the existing tasks.
param(
    [string]$DiscoverDay = "Monday",
    [string]$DiscoverTime = "06:30",
    [string]$LeadsStart = "07:00",
    [int]$LeadsEveryHours = 2,
    [string]$FollowupTime = "07:30",
    [int]$DriveSyncEveryMinutes = 15
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$pythonw = Join-Path $root ".venv\Scripts\pythonw.exe"
if (-not (Test-Path $pythonw)) { throw "Run scripts\setup.ps1 first." }

# Runs as you, only while you're logged on (OneDrive must be running), and catches up
# on runs missed while the PC was off or asleep.
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 3) -MultipleInstances IgnoreNew
# Discover and Follow-up also wake the laptop from sleep at their scheduled time.
$wakeSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 3) -MultipleInstances IgnoreNew -WakeToRun

function Add-AgentTask($name, $arguments, $trigger, $description, $taskSettings = $settings) {
    $action = New-ScheduledTaskAction -Execute $pythonw -Argument "-m origincrest_agent $arguments" `
        -WorkingDirectory $root
    Register-ScheduledTask -TaskPath "\OriginCrest\" -TaskName $name -Action $action -Trigger $trigger `
        -Settings $taskSettings -Description $description -Force | Out-Null
    Write-Host "Scheduled: OriginCrest\$name"
}

$weekly = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DiscoverDay -At $DiscoverTime
Add-AgentTask "1 Discover" "discover" $weekly `
    "Weekly: find new leads for each workbook (Adelaide / other states), then draft RFQs and intros." $wakeSettings

$leads = New-ScheduledTaskTrigger -Daily -At $LeadsStart
$leads.Repetition = (New-ScheduledTaskTrigger -Once -At $LeadsStart `
    -RepetitionInterval (New-TimeSpan -Hours $LeadsEveryHours) -RepetitionDuration (New-TimeSpan -Hours 12)).Repetition
Add-AgentTask "2 Leads" "leads" $leads `
    "Every $LeadsEveryHours h from $($LeadsStart): draft RFQs / intros for new rows in the new-lead tabs."

$weekdays = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At $FollowupTime
Add-AgentTask "3 Follow-up" "followup" $weekdays `
    "Weekdays: check replies, save follow-up drafts, update Last contacted / Next action date." $wakeSettings

# Checks every few minutes, all day; copies only when a file has changed (otherwise it does nothing).
$sync = New-ScheduledTaskTrigger -Daily -At "00:00"
$sync.Repetition = (New-ScheduledTaskTrigger -Once -At "00:00" `
    -RepetitionInterval (New-TimeSpan -Minutes $DriveSyncEveryMinutes) -RepetitionDuration (New-TimeSpan -Hours 24)).Repetition
Add-AgentTask "4 Drive sync" "drive-sync" $sync `
    "Every $DriveSyncEveryMinutes min: if any file in the OneDrive Origincrest folder changed, copy it to Google Drive (never deletes)."

# Waking from sleep also needs "Allow wake timers" in the power plan. Turn it on for when the laptop is
# plugged in only, so it never wakes up in a bag on battery.
powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP RTCWAKE 1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    powercfg /SETACTIVE SCHEME_CURRENT | Out-Null
    Write-Host "Wake timers on (plugged in): 1 Discover and 3 Follow-up wake the laptop from sleep."
} else {
    Write-Host "Couldn't turn on wake timers. Do it in Power Options > Change plan settings >"
    Write-Host "Change advanced power settings > Sleep > Allow wake timers > Plugged in: Enable."
}

Write-Host ""
Write-Host "Done. See them in Task Scheduler > Task Scheduler Library > OriginCrest."
Write-Host "Remove them with: powershell -ExecutionPolicy Bypass -File scripts\uninstall_tasks.ps1"
