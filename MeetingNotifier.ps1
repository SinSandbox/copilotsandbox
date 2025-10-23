# Meeting Notification Monitor
# This script monitors Outlook calendar for upcoming meetings and displays popup notifications
# Designed to run as a scheduled task in the background

# Configuration
$CheckIntervalSeconds = 300  # Check every 5 minutes
$ReminderMinutesBefore = 5   # Notify X minutes before meeting starts
$LogFile = "$env:TEMP\MeetingNotifier.log"

# Function to write log entries
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogFile -Append
}

# Function to show popup notification
function Show-MeetingNotification {
    param(
        [string]$Subject,
        [string]$StartTime,
        [string]$MeetingType
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $notification = New-Object System.Windows.Forms.NotifyIcon
    $notification.Icon = [System.Drawing.SystemIcons]::Information
    $notification.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $notification.BalloonTipTitle = "$MeetingType Meeting Reminder"
    $notification.BalloonTipText = "Meeting: $Subject`nStarts: $StartTime"
    $notification.Visible = $true
    $notification.ShowBalloonTip(10000)
    
    # Keep notification visible
    Start-Sleep -Seconds 10
    $notification.Dispose()
}

# Function to check Outlook meetings
function Check-OutlookMeetings {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $calendar = $namespace.GetDefaultFolder(9) # 9 = olFolderCalendar
        
        $now = Get-Date
        $checkUntil = $now.AddMinutes($ReminderMinutesBefore + 1)
        
        # Filter for appointments in the next X minutes
        $filter = "[Start] >= '$($now.ToString("g"))' AND [Start] <= '$($checkUntil.ToString("g"))'"
        $appointments = $calendar.Items.Restrict($filter)
        $appointments.Sort("[Start]")
        
        foreach ($appt in $appointments) {
            $startTime = [DateTime]$appt.Start
            $minutesUntil = ($startTime - $now).TotalMinutes
            
            # Check if we should notify (within reminder window and not notified yet)
            if ($minutesUntil -le $ReminderMinutesBefore -and $minutesUntil -gt 0) {
                $meetingId = "$($appt.Subject)_$($appt.Start)"
                $notifiedFile = "$env:TEMP\notified_$($meetingId -replace '[^\w]','_').txt"
                
                # Only notify if we haven't already notified for this meeting
                if (-not (Test-Path $notifiedFile)) {
                    $meetingType = if ($appt.Location -match "Microsoft Teams|Teams Meeting|teams.microsoft.com") {
                        "Teams"
                    } else {
                        "Outlook"
                    }
                    
                    Show-MeetingNotification -Subject $appt.Subject -StartTime $startTime.ToString("HH:mm") -MeetingType $meetingType
                    Write-Log "Notification shown for: $($appt.Subject) at $($startTime.ToString("HH:mm"))"
                    
                    # Mark as notified
                    "Notified" | Out-File -FilePath $notifiedFile
                }
            }
        }
        
        # Cleanup old notification markers (older than 24 hours)
        Get-ChildItem "$env:TEMP\notified_*.txt" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-24) } | Remove-Item -Force
        
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
    } catch {
        Write-Log "Error checking Outlook meetings: $_"
    }
}

# Main monitoring loop
Write-Log "Meeting notification monitor started"

while ($true) {
    try {
        Check-OutlookMeetings
        Start-Sleep -Seconds $CheckIntervalSeconds
    } catch {
        Write-Log "Error in main loop: $_"
        Start-Sleep -Seconds 60  # Wait a bit before retrying on error
    }
}
