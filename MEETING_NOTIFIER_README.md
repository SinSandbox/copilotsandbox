# Meeting Notifier

PowerShell scripts to monitor Outlook and Teams meetings and display popup notifications.

## Features

- ✅ Monitors Outlook calendar for upcoming meetings
- ✅ Detects both Outlook and Teams meetings
- ✅ Shows balloon notifications 5 minutes before meetings
- ✅ Runs silently in the background
- ✅ Doesn't disrupt other desktop workloads
- ✅ Prevents duplicate notifications
- ✅ Can run as a Windows scheduled task
- ✅ Logs all activity for troubleshooting

## Files

- **MeetingNotifier.ps1** - Main monitoring script
- **Setup-MeetingNotifier.ps1** - Automated task scheduler setup

## Quick Setup

### Option 1: Automated Setup (Recommended)

1. Open PowerShell **as Administrator**
2. Navigate to the script directory:
   ```powershell
   cd "c:\Users\mosin\source\repos\copilotsandbox"
   ```
3. Run the setup script:
   ```powershell
   .\Setup-MeetingNotifier.ps1
   ```
4. Start the task:
   ```powershell
   Start-ScheduledTask -TaskName "MeetingNotifier"
   ```

### Option 2: Manual Setup

1. Open Task Scheduler (`taskschd.msc`)
2. Click **Create Task**
3. **General tab:**
   - Name: `MeetingNotifier`
   - Run whether user is logged on or not: ❌ (unchecked)
   - Run with highest privileges: ✅ (checked)
4. **Triggers tab:**
   - New → At log on → Specific user
5. **Actions tab:**
   - New → Start a program
   - Program: `powershell.exe`
   - Arguments: `-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "c:\Users\mosin\source\repos\copilotsandbox\MeetingNotifier.ps1"`
6. **Settings tab:**
   - Allow task to be run on demand: ✅
   - Stop task if it runs longer than: ❌ (unchecked)
   - If running task doesn't end when requested: Do not stop
7. Click **OK** to save

## Configuration

Edit `MeetingNotifier.ps1` to customize:

```powershell
$CheckIntervalSeconds = 300  # How often to check (default: 5 minutes)
$ReminderMinutesBefore = 5   # When to notify (default: 5 minutes before)
```

## Logs

View activity logs at:
```
%TEMP%\MeetingNotifier.log
```

Or open in PowerShell:
```powershell
notepad $env:TEMP\MeetingNotifier.log
```

## Management Commands

### Start the notifier
```powershell
Start-ScheduledTask -TaskName "MeetingNotifier"
```

### Stop the notifier
```powershell
Stop-ScheduledTask -TaskName "MeetingNotifier"
```

### Check status
```powershell
Get-ScheduledTask -TaskName "MeetingNotifier" | Select-Object State, LastRunTime, NextRunTime
```

### Remove the task
```powershell
Unregister-ScheduledTask -TaskName "MeetingNotifier" -Confirm:$false
```

## Troubleshooting

### No notifications appearing

1. Check if task is running:
   ```powershell
   Get-ScheduledTask -TaskName "MeetingNotifier"
   ```
2. Check the log file for errors
3. Ensure Outlook is installed and configured
4. Verify you have upcoming meetings in Outlook calendar

### Script execution policy error

Run this in PowerShell as Administrator:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Outlook COM object errors

1. Ensure Outlook is installed
2. Try opening Outlook manually first
3. Check if Outlook profile is configured

## Requirements

- Windows 10/11
- Microsoft Outlook (desktop version)
- PowerShell 5.1 or later
- User must be logged in for notifications to display

## Notes

- The script uses Outlook's COM interface to read calendar data
- Teams meetings are detected by checking the meeting location field
- Notification markers are stored in `%TEMP%` and cleaned up after 24 hours
- The script runs continuously but uses minimal resources (checks every 5 minutes)
- Balloon notifications automatically dismiss after 10 seconds
