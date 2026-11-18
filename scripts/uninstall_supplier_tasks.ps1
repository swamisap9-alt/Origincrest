# Removes the supplier agent's scheduled tasks (folder "OriginCrest Suppliers"). The outreach tasks are kept.
Get-ScheduledTask -TaskPath "\OriginCrest Suppliers\" -ErrorAction SilentlyContinue |
    Unregister-ScheduledTask -Confirm:$false
Write-Host "OriginCrest supplier tasks removed."
