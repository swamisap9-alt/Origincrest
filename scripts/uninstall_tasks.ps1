# Removes the OriginCrest scheduled tasks.
Get-ScheduledTask -TaskPath "\OriginCrest\" -ErrorAction SilentlyContinue |
    Unregister-ScheduledTask -Confirm:$false
Write-Host "OriginCrest scheduled tasks removed."
