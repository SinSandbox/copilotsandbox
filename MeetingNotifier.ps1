# Meeting Notification Monitor
# This script monitors Outlook calendar for upcoming meetings and displays popup notifications
# Works with the new Outlook for Windows app using Microsoft Graph API
# Designed to run as a scheduled task in the background

# Configuration
$CheckIntervalSeconds = 300  # Check every 5 minutes
$ReminderMinutesBefore = 5   # Notify X minutes before meeting starts
$LogFile = "$env:TEMP\MeetingNotifier.log"
$OutlookDbPath = "$env:LOCALAPPDATA\Microsoft\Olk\store.db"  # New Outlook SQLite database

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

# Function to get new Outlook database path
function Get-OutlookDatabasePath {
    # Try multiple possible locations for the new Outlook database
    $possiblePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Olk\store.db",
        "$env:LOCALAPPDATA\Microsoft\Outlook\store.db",
        "$env:LOCALAPPDATA\Packages\Microsoft.OutlookForWindows_8wekyb3d8bbwe\LocalCache\Local\Olk\store.db"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return $path
        }
    }
    
    return $null
}

# Function to check Outlook meetings using new Outlook database
function Check-OutlookMeetings {
    try {
        # First, try the new Outlook approach
        $dbPath = Get-OutlookDatabasePath
        
        if ($dbPath) {
            Check-NewOutlookMeetings -DbPath $dbPath
        } else {
            # Fallback to classic Outlook COM if new Outlook not found
            Write-Log "New Outlook database not found, attempting classic Outlook COM..."
            Check-ClassicOutlookMeetings
        }
        
        # Cleanup old notification markers (older than 24 hours)
        Get-ChildItem "$env:TEMP\notified_*.txt" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-24) } | Remove-Item -Force
        
    } catch {
        Write-Log "Error checking Outlook meetings: $_"
    }
}

# Function to check new Outlook meetings via local cache
function Check-NewOutlookMeetings {
    param([string]$DbPath)
    
    try {
        # Load System.Data.SQLite if available, otherwise use workaround
        $sqliteLoaded = $false
        try {
            Add-Type -Path "$env:ProgramFiles\System.Data.SQLite\netstandard2.0\System.Data.SQLite.dll" -ErrorAction SilentlyContinue
            $sqliteLoaded = $true
        } catch {
            # SQLite not available, try alternative approach via Windows Calendar API
        }
        
        if (-not $sqliteLoaded) {
            Write-Log "SQLite not available. Using Windows.ApplicationModel.Appointments API..."
            Check-WindowsCalendarAPI
            return
        }
        
        $connection = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$DbPath;Version=3;Read Only=True;")
        $connection.Open()
        
        $now = Get-Date
        $checkUntil = $now.AddMinutes($ReminderMinutesBefore + 1)
        
        $query = @"
SELECT Subject, Location, Start, End 
FROM CalendarItems 
WHERE Start >= @now AND Start <= @checkUntil
ORDER BY Start
"@
        
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $command.Parameters.AddWithValue("@now", $now.ToString("o")) | Out-Null
        $command.Parameters.AddWithValue("@checkUntil", $checkUntil.ToString("o")) | Out-Null
        
        $reader = $command.ExecuteReader()
        
        while ($reader.Read()) {
            $subject = $reader["Subject"]
            $location = $reader["Location"]
            $startTime = [DateTime]::Parse($reader["Start"])
            
            $minutesUntil = ($startTime - $now).TotalMinutes
            
            if ($minutesUntil -le $ReminderMinutesBefore -and $minutesUntil -gt 0) {
                $meetingId = "$($subject)_$($startTime.ToString("yyyyMMddHHmm"))"
                $notifiedFile = "$env:TEMP\notified_$($meetingId -replace '[^\w]','_').txt"
                
                if (-not (Test-Path $notifiedFile)) {
                    $meetingType = if ($location -match "Microsoft Teams|Teams Meeting|teams.microsoft.com") {
                        "Teams"
                    } else {
                        "Outlook"
                    }
                    
                    Show-MeetingNotification -Subject $subject -StartTime $startTime.ToString("HH:mm") -MeetingType $meetingType
                    Write-Log "Notification shown for: $subject at $($startTime.ToString("HH:mm"))"
                    
                    "Notified" | Out-File -FilePath $notifiedFile
                }
            }
        }
        
        $reader.Close()
        $connection.Close()
        
    } catch {
        Write-Log "Error checking new Outlook meetings: $_"
        # Fallback to classic Outlook
        Check-ClassicOutlookMeetings
    }
}

# Function to check Windows Calendar API (works with new Outlook)
function Check-WindowsCalendarAPI {
    try {
        [Windows.ApplicationModel.Appointments.AppointmentManager, Windows.ApplicationModel.Appointments, ContentType = WindowsRuntime] | Out-Null
        [Windows.ApplicationModel.Appointments.AppointmentStore, Windows.ApplicationModel.Appointments, ContentType = WindowsRuntime] | Out-Null
        
        $now = Get-Date
        $checkUntil = $now.AddMinutes($ReminderMinutesBefore + 1)
        
        # Get appointment store
        $storeTask = [Windows.ApplicationModel.Appointments.AppointmentManager]::RequestStoreAsync([Windows.ApplicationModel.Appointments.AppointmentStoreAccessType]::AppCalendarsReadWrite)
        $storeTask.AsTask().Wait()
        $store = $storeTask.GetResults()
        
        if ($null -eq $store) {
            Write-Log "Could not access Windows Calendar API"
            return
        }
        
        # Find appointments
        $findTask = $store.FindAppointmentsAsync($now, ($checkUntil - $now))
        $findTask.AsTask().Wait()
        $appointments = $findTask.GetResults()
        
        foreach ($appt in $appointments) {
            $startTime = $appt.StartTime.DateTime
            $minutesUntil = ($startTime - $now).TotalMinutes
            
            if ($minutesUntil -le $ReminderMinutesBefore -and $minutesUntil -gt 0) {
                $meetingId = "$($appt.Subject)_$($startTime.ToString("yyyyMMddHHmm"))"
                $notifiedFile = "$env:TEMP\notified_$($meetingId -replace '[^\w]','_').txt"
                
                if (-not (Test-Path $notifiedFile)) {
                    $meetingType = if ($appt.Location -match "Microsoft Teams|Teams Meeting|teams.microsoft.com") {
                        "Teams"
                    } else {
                        "Outlook"
                    }
                    
                    Show-MeetingNotification -Subject $appt.Subject -StartTime $startTime.ToString("HH:mm") -MeetingType $meetingType
                    Write-Log "Notification shown for: $($appt.Subject) at $($startTime.ToString("HH:mm"))"
                    
                    "Notified" | Out-File -FilePath $notifiedFile
                }
            }
        }
        
    } catch {
        Write-Log "Error using Windows Calendar API: $_"
    }
}

# Function to check classic Outlook meetings (fallback)
function Check-ClassicOutlookMeetings {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $calendar = $namespace.GetDefaultFolder(9) # 9 = olFolderCalendar
        
        $now = Get-Date
        $checkUntil = $now.AddMinutes($ReminderMinutesBefore + 1)
        
        $filter = "[Start] >= '$($now.ToString("g"))' AND [Start] <= '$($checkUntil.ToString("g"))'"
        $appointments = $calendar.Items.Restrict($filter)
        $appointments.Sort("[Start]")
        
        foreach ($appt in $appointments) {
            $startTime = [DateTime]$appt.Start
            $minutesUntil = ($startTime - $now).TotalMinutes
            
            if ($minutesUntil -le $ReminderMinutesBefore -and $minutesUntil -gt 0) {
                $meetingId = "$($appt.Subject)_$($appt.Start)"
                $notifiedFile = "$env:TEMP\notified_$($meetingId -replace '[^\w]','_').txt"
                
                if (-not (Test-Path $notifiedFile)) {
                    $meetingType = if ($appt.Location -match "Microsoft Teams|Teams Meeting|teams.microsoft.com") {
                        "Teams"
                    } else {
                        "Outlook"
                    }
                    
                    Show-MeetingNotification -Subject $appt.Subject -StartTime $startTime.ToString("HH:mm") -MeetingType $meetingType
                    Write-Log "Notification shown for: $($appt.Subject) at $($startTime.ToString("HH:mm"))"
                    
                    "Notified" | Out-File -FilePath $notifiedFile
                }
            }
        }
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
    } catch {
        Write-Log "Error checking classic Outlook meetings: $_"
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
