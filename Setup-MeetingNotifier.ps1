# Task Scheduler Setup Script
# Run this script with Administrator privileges to configure the meeting notifier as a scheduled task

$TaskName = "MeetingNotifier"
$ScriptPath = "$PSScriptRoot\MeetingNotifier.ps1"
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

Write-Host "Setting up Meeting Notifier Scheduled Task..." -ForegroundColor Green

# Check if task already exists
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Task already exists. Removing old task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# Create task action - run PowerShell hidden
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""

# Create task trigger - at logon
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $CurrentUser

# Create task settings - allow running on battery, don't stop if going on battery
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -DontStopOnIdleEnd `
    -ExecutionTimeLimit (New-TimeSpan -Days 0)  # No time limit

# Create task principal - run as current user with highest privileges
$principal = New-ScheduledTaskPrincipal -UserId $CurrentUser -LogonType Interactive -RunLevel Highest

# Register the task
Register-ScheduledTask -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Monitors Outlook and Teams meetings and displays popup notifications"

Write-Host "`nTask '$TaskName' has been created successfully!" -ForegroundColor Green
Write-Host "The task will:" -ForegroundColor Cyan
Write-Host "  - Start automatically when you log in" -ForegroundColor White
Write-Host "  - Run hidden in the background" -ForegroundColor White
Write-Host "  - Check for meetings every 5 minutes" -ForegroundColor White
Write-Host "  - Notify 5 minutes before each meeting" -ForegroundColor White
Write-Host "`nTo start the task now, run:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor White
Write-Host "`nTo view logs, check:" -ForegroundColor Yellow
Write-Host "  $env:TEMP\MeetingNotifier.log" -ForegroundColor White
Write-Host "`nTo stop the task, run:" -ForegroundColor Yellow
Write-Host "  Stop-ScheduledTask -TaskName '$TaskName'" -ForegroundColor White
Write-Host "`nTo remove the task, run:" -ForegroundColor Yellow
Write-Host "  Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false" -ForegroundColor White
