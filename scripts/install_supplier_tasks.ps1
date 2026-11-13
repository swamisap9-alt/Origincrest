# Registers the supplier agent's scheduled tasks in Task Scheduler, in their own folder
# ("OriginCrest Suppliers"), so the outreach agent's install / uninstall scripts never touch them.
#   powershell -ExecutionPolicy Bypass -File scripts\install_supplier_tasks.ps1
# Re-run any time to change times; it only replaces the supplier tasks.
param(
    [int]$SearchEveryDays = 3,
    [string]$SearchTime = "06:00",
    [string]$FollowupTime = "07:45"
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$pythonw = Join-Path $root ".venv\Scripts\pythonw.exe"
if (-not (Test-Path $pythonw)) { throw "Run scripts\setup.ps1 first." }
if (-not (Test-Path (Join-Path $root "config.suppliers.yaml"))) {
    throw "config.suppliers.yaml not found. Create it: Copy-Item config.suppliers.example.yaml config.suppliers.yaml"
}

# An earlier version registered the supplier search as "OriginCrest\5 Suppliers"; remove it so it can't run twice.
$old = Get-ScheduledTask -TaskPath "\OriginCrest\" -TaskName "5 Suppliers" -ErrorAction SilentlyContinue
if ($old) {
    $old | Unregister-ScheduledTask -Confirm:$false
    Write-Host "Removed the old task OriginCrest\5 Suppliers"
}

# Runs as you while you're logged on, catches up on runs missed while the PC was off, and wakes it when plugged in.
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 3) -MultipleInstances IgnoreNew -WakeToRun

function Add-SupplierTask($name, $arguments, $trigger, $description) {
    $action = New-ScheduledTaskAction -Execute $pythonw -Argument "-m origincrest_agent.supplier_agent $arguments" `
        -WorkingDirectory $root
    Register-ScheduledTask -TaskPath "\OriginCrest Suppliers\" -TaskName $name -Action $action -Trigger $trigger `
        -Settings $settings -Description $description -Force | Out-Null
    Write-Host "Scheduled: OriginCrest Suppliers\$name"
}

$search = New-ScheduledTaskTrigger -Daily -DaysInterval $SearchEveryDays -At $SearchTime
Add-SupplierTask "1 Supplier search" "search" $search `
    "Every $SearchEveryDays days: find Indonesian coffee suppliers, compare them with the current vendor, draft sourcing enquiries."

$weekdays = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At $FollowupTime
Add-SupplierTask "2 Supplier follow-up" "followup" $weekdays `
    "Weekdays: check supplier replies and draft follow-ups to enquiries unanswered after the set number of days."

Write-Host ""
Write-Host "Done. See them in Task Scheduler > Task Scheduler Library > OriginCrest Suppliers."
Write-Host "Waking from sleep needs 'Allow wake timers' (Power Options); the outreach installer turns it on."
Write-Host "Remove them with: powershell -ExecutionPolicy Bypass -File scripts\uninstall_supplier_tasks.ps1"
