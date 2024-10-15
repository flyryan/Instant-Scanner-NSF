# =====================================
# Trend Micro ScanMail for Domino Instant NSF Scanner Script
# © Trend Micro, Inc. 2024
# =====================================

# =====================================
#          Universal Settings
# =====================================

$Folder = "C:\path\to\your\NSF\directory"       # Directory to monitor for new NSF files.
$ScanmailPath = "C:\path\to\domino\cconsole.exe" # Path to the ScanMail console executable.
$MaxConcurrentScans = 4  # Maximum number of concurrent scans to run simultaneously.

# =====================================
#          Configuration Toggles
# =====================================

# Toggle for Logging
$EnableLogging = $false
# Description: Enables logging of script activity to a file for auditing and troubleshooting.

# Logging Configuration
if ($EnableLogging) {
    $LogFilePath = "C:\path\to\logfile.log"  # Path to the log file where script activity will be recorded.
}

# Toggle for Email Notifications on Errors
$EnableEmailNotification = $false
# Description: Sends an email notification whenever an error occurs during processing. Requires SMTP credential storage (see installation guide)

# Email Notification Configuration
if ($EnableEmailNotification) {
    # Retrieve credentials from Credential Manager
    $Credential = Get-StoredCredential -Target "SMTPCredential"

    $EmailSettings = @{
        To         = "your.email@domain.com"          # Recipient email address.
        From       = "script.alerts@domain.com"       # Sender email address.
        Subject    = "ScanMail Script Error Notification"  # Email subject line.
        SmtpServer = "smtp.yourdomain.com"            # SMTP server address.
        Credential = $Credential                      # Credentials for SMTP server authentication.
        UseSsl     = $true                            # Enable SSL for SMTP connection if required.
        Port       = 587                              # SMTP port if different from default (25).
    }
}

# Toggle for Heartbeat Mechanism
$EnableHeartbeat = $false
# Description: Updates a heartbeat file periodically to indicate that the script is running.

# Heartbeat Mechanism Configuration
if ($EnableHeartbeat) {
    $HeartbeatFile = "C:\path\to\heartbeat.txt"   # Path to the heartbeat file to indicate the script is running.
    $HeartbeatInterval = 60000  # Interval in milliseconds for updating the heartbeat file (e.g., 60000ms = 60 seconds).
}

# Toggle for Self-Check Mechanism
$EnableSelfCheck = $true
# Description: Periodically checks worker jobs and restarts them if they have stopped running.

# Self-Check Mechanism Configuration
if ($EnableSelfCheck) {
    $SelfCheckInterval = 300000  # Interval in milliseconds for checking worker jobs (e.g., 300000ms = 5 minutes).
}

# Toggle for Processed Files Logging
$EnableProcessedFilesLog = $false
# Description: Logs each successfully processed file to a log file for record-keeping.

# Processed Files Log Configuration
if ($EnableProcessedFilesLog) {
    $ProcessedFilesLog = "C:\path\to\processedfiles.log"  # Path to the log file where processed files are recorded.
}

# =====================================
#            Helper Functions
# =====================================

# Logging Function
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    # Write to console
    Write-Host $LogEntry
    # Append to log file if logging is enabled
    if ($EnableLogging) {
        Add-Content -Path $LogFilePath -Value $LogEntry
    }
}

# Email Notification Function
function Send-ErrorEmail {
    param(
        [string]$ErrorMessage
    )
    if ($EnableEmailNotification) {
        $EmailParams = $EmailSettings.Clone()
        $EmailParams.Subject = "$($EmailSettings.Subject): Error Occurred"
        $EmailParams.Body    = $ErrorMessage
        try {
            Send-MailMessage @EmailParams
            Write-Log "Error email sent successfully."
        } catch {
            Write-Log "Failed to send error email. Error: $_" "ERROR"
        }
    }
}

# Heartbeat Update Function
function Update-Heartbeat {
    if ($EnableHeartbeat) {
        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Set-Content -Path $HeartbeatFile -Value $Timestamp
        Write-Log "Heartbeat updated."
    }
}

# Processed Files Logging Function
function Log-ProcessedFile {
    param(
        [string]$File
    )
    if ($EnableProcessedFilesLog) {
        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogEntry = "$Timestamp,$File"
        Add-Content -Path $ProcessedFilesLog -Value $LogEntry
        Write-Log "Logged processed file: $File"
    }
}

# =====================================
#          Main Script Logic
# =====================================

# Create a concurrent queue
$Queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# Function to process the queue
$ProcessQueueScript = {
    param(
        $Queue,
        $ScanmailPath,
        $EnableLogging,
        $LogFilePath,
        $EnableEmailNotification,
        $EmailSettings,
        $EnableHeartbeat,
        $HeartbeatFile,
        $EnableProcessedFilesLog,
        $ProcessedFilesLog
    )

    # Define helper functions within the job scope
    function Write-Log {
        param(
            [string]$Message,
            [string]$Level = "INFO"
        )
        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogEntry = "[$Timestamp] [$Level] $Message"
        Write-Host $LogEntry
        if ($EnableLogging) {
            Add-Content -Path $LogFilePath -Value $LogEntry
        }
    }

    function Send-ErrorEmail {
        param(
            [string]$ErrorMessage
        )
        if ($EnableEmailNotification) {
            $EmailParams = $EmailSettings.Clone()
            $EmailParams.Subject = "$($EmailSettings.Subject): Error Occurred"
            $EmailParams.Body    = $ErrorMessage
            try {
                Send-MailMessage @EmailParams
                Write-Log "Error email sent successfully."
            } catch {
                Write-Log "Failed to send error email. Error: $_" "ERROR"
            }
        }
    }

    function Update-Heartbeat {
        if ($EnableHeartbeat) {
            $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Set-Content -Path $HeartbeatFile -Value $Timestamp
            Write-Log "Heartbeat updated."
        }
    }

    function Log-ProcessedFile {
        param(
            [string]$File
        )
        if ($EnableProcessedFilesLog) {
            $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $LogEntry = "$Timestamp,$File"
            Add-Content -Path $ProcessedFilesLog -Value $LogEntry
            Write-Log "Logged processed file: $File"
        }
    }

    while ($true) {
        if ($Queue.TryDequeue([ref]$File)) {
            # Wait until the file is fully accessible
            while ($true) {
                try {
                    $Stream = [System.IO.File]::Open($File, 'Open', 'ReadWrite', 'None')
                    $Stream.Close()
                    break
                } catch {
                    Start-Sleep -Milliseconds 500
                }
            }

            # Properly quote the file path
            $QuotedFilePath = '"' + $File + '"'
            $ScanmailCommand = "-c 'load SMDdbs -manual $QuotedFilePath'"

            try {
                # Execute the scan command
                & $ScanmailPath $ScanmailCommand
                Write-Log "Scanned NSF file: $File"
                # Log the processed file
                Log-ProcessedFile $File
            } catch {
                $ErrorMessage = "Failed to scan $File. Error: $_"
                Write-Log $ErrorMessage "ERROR"
                Send-ErrorEmail $ErrorMessage
            }
        } else {
            # No items in queue, wait a bit
            Start-Sleep -Milliseconds 500
        }

        # Update heartbeat in worker thread
        Update-Heartbeat
    }
}

# Start worker jobs
for ($i = 1; $i -le $MaxConcurrentScans; $i++) {
    Start-Job -ScriptBlock $ProcessQueueScript -ArgumentList $Queue, $ScanmailPath, $EnableLogging, $LogFilePath, $EnableEmailNotification, $EmailSettings, $EnableHeartbeat, $HeartbeatFile, $EnableProcessedFilesLog, $ProcessedFilesLog
}

# Set up the FileSystemWatcher
$Watcher = New-Object System.IO.FileSystemWatcher
$Watcher.Path = $Folder
$Watcher.Filter = "*.nsf"
$Watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName
$Watcher.IncludeSubdirectories = $true
$Watcher.EnableRaisingEvents = $true

$OnCreatedAction = {
    param($SourceEventArgs)
    $NewFile = $SourceEventArgs.FullPath
    $Queue.Enqueue($NewFile)
    Write-Log "Enqueued new NSF file: $NewFile"
}

Register-ObjectEvent $Watcher Created -Action $OnCreatedAction

# =====================================
#         Heartbeat Mechanism
# =====================================

if ($EnableHeartbeat) {
    Write-Log "Heartbeat mechanism enabled."
    $HeartbeatTimer = [System.Timers.Timer]::new($HeartbeatInterval)
    $HeartbeatTimer.AutoReset = $true
    $HeartbeatTimer.Enabled = $true
    $HeartbeatTimer.Add_Elapsed({
        Update-Heartbeat
    })
}

# =====================================
#         Self-Check Mechanism
# =====================================

if ($EnableSelfCheck) {
    Write-Log "Self-check mechanism enabled."
    $CheckJobsTimer = [System.Timers.Timer]::new($SelfCheckInterval)
    $CheckJobsTimer.AutoReset = $true
    $CheckJobsTimer.Enabled = $true
    $CheckJobsTimer.Add_Elapsed({
        # Ensure Write-Log is available in this scope
        function Write-Log {
            param(
                [string]$Message,
                [string]$Level = "INFO"
            )
            $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $LogEntry = "[$Timestamp] [$Level] $Message"
            Write-Host $LogEntry
            if ($using:EnableLogging) {
                Add-Content -Path $using:LogFilePath -Value $LogEntry
            }
        }

        $Jobs = Get-Job
        foreach ($Job in $Jobs) {
            if ($Job.State -ne "Running") {
                $WarningMessage = "Job ID $($Job.Id) is not running. Restarting..."
                Write-Log $WarningMessage "WARNING"
                Remove-Job $Job
                # Restart the job
                Start-Job -ScriptBlock $using:ProcessQueueScript -ArgumentList $using:Queue, $using:ScanmailPath, $using:EnableLogging, $using:LogFilePath, $using:EnableEmailNotification, $using:EmailSettings, $using:EnableHeartbeat, $using:HeartbeatFile, $using:EnableProcessedFilesLog, $using:ProcessedFilesLog
            }
        }
    })
}

# =====================================
#         Script Execution Start
# =====================================

Write-Log "Monitoring directory: $Folder"
Write-Log "Script started successfully."

# Keep the script running
Wait-Event

# =====================================
#         Exception Handling
# =====================================

# Global exception handler
$ErrorActionPreference = 'Stop'
trap {
    $ErrorMessage = "A fatal error occurred: $_"
    Write-Log $ErrorMessage "ERROR"
    Send-ErrorEmail $ErrorMessage
    exit 1
}