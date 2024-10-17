# =====================================
# Trend Micro ScanMail for Domino Instant NSF Scanner Script
# © Trend Micro, Inc. 2024
# =====================================

# =====================================
#          Universal Settings
# =====================================

$Folder = "C:\path\to\nsf\directory"               # Directory to monitor for new NSF files.
$ConsolePath = "C:\path\to\Domino\nserver.exe"      # Path to the Domino server console executable.
$MaxConcurrentScans = 4                              # Maximum number of concurrent scans to run simultaneously.
$ConsoleLogLevel = "INFO"                            # Valid options: DEBUG, INFO, WARNING, ERROR

# =====================================
#          Configuration Toggles
# =====================================

# Toggle for Logging
$EnableLogging = $false
$FileLogLevel = "INFO"                                # Valid options: DEBUG, INFO, WARNING, ERROR
# Description: Enables logging of script activity to a file for auditing and troubleshooting.

# Logging Configuration
if ($EnableLogging) {
    $LogFilePath = "C:\path\to\logfile.log"           # Path to the log file where script activity will be recorded.
}

# Toggle for Email Notifications on Errors
$EnableEmailNotification = $false
# Description: Sends an email notification whenever an error occurs during processing.

# Email Notification Configuration
if ($EnableEmailNotification) {
    # Retrieve credentials from Credential Manager
    # Note: You need to have stored credentials in Credential Manager with the target "SMTPCredential"
    $Credential = Get-StoredCredential -Target "SMTPCredential"

    $EmailSettings = @{
        To         = "your.email@domain.com"                 # Recipient email address.
        From       = "script.alerts@domain.com"              # Sender email address.
        Subject    = "ScanMail Script Error Notification"    # Email subject line.
        SmtpServer = "smtp.yourdomain.com"                   # SMTP server address.
        Credential = $Credential                             # Credentials for SMTP server authentication.
        UseSsl     = $true                                   # Enable SSL for SMTP connection if required.
        Port       = 587                                     # SMTP port if different from default (25).
    }
}

# Toggle for Heartbeat Mechanism
$EnableHeartbeat = $false
# Description: Updates a heartbeat file periodically to indicate that the script is running.

# Heartbeat Mechanism Configuration
if ($EnableHeartbeat) {
    $HeartbeatFile = "C:\path\to\heartbeat.txt"          # Path to the heartbeat file to indicate the script is running.
    $HeartbeatInterval = 60000                           # Interval in milliseconds for updating the heartbeat file (e.g., 60000ms = 60 seconds).
}

# Toggle for Processed Files Logging
$EnableProcessedFilesLog = $false
# Description: Logs each processed file to a log file for record-keeping.

# Processed Files Log Configuration
if ($EnableProcessedFilesLog) {
    $ProcessedFilesLog = "C:\path\to\processedfiles.log"  # Path to the log file where processed files are recorded.
}

# Toggle for Self-Check Mechanism
$EnableSelfCheck = $true
# Description: Periodically checks worker jobs and restarts them if they have stopped running.

# Self-Check Mechanism Configuration
if ($EnableSelfCheck) {
    $SelfCheckInterval = 300000                          # Interval in milliseconds for checking worker jobs (e.g., 300000ms = 5 minutes).
}

# =====================================
#          Global Variables
# =====================================

# Ensure ErrorActionPreference is set to stop on errors
$ErrorActionPreference = 'Stop'

# Import ThreadJob module
Import-Module ThreadJob -ErrorAction Stop

# Logging Levels Map
$Global:LogLevels = @{
    "DEBUG"   = 1
    "INFO"    = 2
    "WARNING" = 3
    "ERROR"   = 4
}

# Create a concurrent queue for log messages
$Global:LogQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()

# Create a concurrent queue for file processing
$Global:Queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# Initialize a hash table to track recently processed files with timestamps
$Global:ProcessedFiles = @{}

# Array to track active worker jobs
$Global:ActiveWorkerJobs = @()

# =====================================
#            Helper Functions
# =====================================

# Global Logging Function
function global:Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"

    # Enqueue the log message
    $Global:LogQueue.Enqueue(@($LogEntry, $Level))

    # Only write to console if message level is >= ConsoleLogLevel
    if ($Global:LogLevels[$Level] -ge $Global:LogLevels[$ConsoleLogLevel]) {
        # Optionally, write to the console with different colors based on level
        switch ($Level) {
            "INFO"    { Write-Host $LogEntry -ForegroundColor Green }
            "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
            "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
            "DEBUG"   { Write-Host $LogEntry -ForegroundColor Cyan }
            default   { Write-Host $LogEntry }
        }
    }
}

# Global Email Notification Function
function global:Send-ErrorEmail {
    param(
        [string]$ErrorMessage
    )
    if ($EnableEmailNotification) {
        # Clone the EmailSettings to avoid modifying the original object
        $ClonedEmailSettings = $EmailSettings.Clone()

        $ClonedEmailSettings.Subject = "$($ClonedEmailSettings.Subject): Error Occurred"
        $ClonedEmailSettings.Body    = $ErrorMessage
        try {
            Send-MailMessage @ClonedEmailSettings
            global:Write-Log "Error email sent successfully." "INFO"
        } catch {
            global:Write-Log "Failed to send error email. Error: $_" "ERROR"
        }
    }
}

# Global Cleanup Function
function global:Cleanup-Script {
    global:Write-Log "Cleaning up background jobs and event subscriptions." "INFO"

    try {
        # Stop and remove all relevant jobs
        $JobsToCleanup = Get-Job | Where-Object { 
            $_.Name -like 'WorkerJob_*' -or 
            $_.Name -eq 'LogProcessorJob' -or 
            $_.Name -eq 'HeartbeatJob' -or 
            $_.Name -eq 'FileCreatedEvent' -or 
            $_.Name -eq 'SelfCheckJob' 
        }

        foreach ($Job in $JobsToCleanup) {
            global:Write-Log "Stopping job: $($Job.Name) (ID: $($Job.Id))" "DEBUG"
            Stop-Job -Job $Job -ErrorAction SilentlyContinue
            Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
            global:Write-Log "Stopped and removed job: $($Job.Name)" "DEBUG"
        }

        # Unregister all event subscribers with valid SubscriptionId
        global:Write-Log "Unregistering event subscribers." "DEBUG"
        Get-EventSubscriber | ForEach-Object {
            if ($_.Id) {
                try {
                    Unregister-Event -SubscriptionId $_.Id -Force -ErrorAction SilentlyContinue
                    global:Write-Log "Unregistered event subscriber: $($_.Id)" "DEBUG"
                } catch {
                    global:Write-Log "Failed to unregister event subscriber: $_" "ERROR"
                }
            } else {
                global:Write-Log "Skipped unregistering an event subscriber with null SubscriptionId." "DEBUG"
            }
        }

        # Unregister FileSystemWatcher events if any
        if ($Watcher) {
            global:Write-Log "Unregistering FileSystemWatcher events." "DEBUG"
            try {
                Unregister-Event -SourceIdentifier FileCreatedEvent -ErrorAction SilentlyContinue
                global:Write-Log "Unregistered FileSystemWatcher events." "DEBUG"
            } catch {
                global:Write-Log "Failed to unregister FileSystemWatcher events: $_" "ERROR"
            }
        }

        # Remove global variables
        global:Write-Log "Removing global variables." "DEBUG"
        $globalVariables = @("Queue", "LogQueue", "HeartbeatJob", "SelfCheckJob", "ProcessedFiles", "ActiveWorkerJobs")
        foreach ($var in $globalVariables) {
            if (Get-Variable -Name $var -Scope Global -ErrorAction SilentlyContinue) {
                Remove-Variable -Name $var -Scope Global -ErrorAction SilentlyContinue
                global:Write-Log "Removed global variable: $var" "DEBUG"
            }
        }

        global:Write-Log "Cleanup completed successfully." "INFO"
    } catch {
        global:Write-Log "An error occurred during cleanup: $_" "ERROR"
    }
}

# =====================================
#         Initial Cleanup on Start
# =====================================

# Global Initial Cleanup Function
function global:Initial-Cleanup {
    global:Write-Log "Performing initial cleanup of existing jobs..." "INFO"

    # Get all jobs that match the naming convention
    $ExistingJobs = Get-Job | Where-Object { 
        $_.Name -like 'WorkerJob_*' -or 
        $_.Name -eq 'LogProcessorJob' -or 
        $_.Name -eq 'HeartbeatJob' -or 
        $_.Name -eq 'Powershell.Exi' -or
        $_.Name -eq 'SelfCheckJob' 
    }

    if ($ExistingJobs) {
        foreach ($Job in $ExistingJobs) {
            global:Write-Log "Stopping and removing existing job: $($Job.Name) (ID: $($Job.Id))" "WARNING"
            Stop-Job $Job -ErrorAction SilentlyContinue
            Remove-Job $Job -Force -ErrorAction SilentlyContinue
            global:Write-Log "Stopped and removed job: $($Job.Name)" "DEBUG"
        }
        global:Write-Log "Initial cleanup completed." "INFO"
    } else {
        global:Write-Log "No existing jobs found. Proceeding..." "INFO"
    }
}

# Call the initial cleanup function
global:Initial-Cleanup

# =====================================
#         Exception Handling
# =====================================

# Global Exception Handler
trap {
    $ErrorMessage = "A fatal error occurred: $_"
    global:Write-Log $ErrorMessage "ERROR"
    global:Send-ErrorEmail $ErrorMessage
    global:Cleanup-Script
    exit 1
}

# Register the cleanup function to run on script exit
Register-EngineEvent PowerShell.Exiting -Action { global:Cleanup-Script } | Out-Null

# Handle Ctrl+C and other console interruptions if supported
$consoleType = [Console].GetType()
$cancelKeyPressEvent = $consoleType.GetEvent("CancelKeyPress")

if ($cancelKeyPressEvent) {
    $Handler = {
        global:Cleanup-Script
        exit
    }
    try {
        $cancelKeyPressEvent.AddEventHandler([Console]::class, $Handler)
        global:Write-Log "CancelKeyPress handler registered successfully." "INFO"
    } catch {
        global:Write-Log "Failed to register CancelKeyPress handler: $_" "ERROR"
    }
} else {
    global:Write-Log "CancelKeyPress event is not available in this PowerShell host." "DEBUG"
}

# =====================================
#          Main Script Logic
# =====================================

try {
    # global:Write-Log "Monitoring directory: $Folder" "INFO"
    global:Write-Log "Script initializing." "INFO"

    # Function to start a worker job
    function Start-WorkerJob {
        param(
            [string]$FilePath
        )

        # Start a new thread job to process the file
        $Job = Start-ThreadJob -ScriptBlock {
            param(
                $ConsolePath,
                $FilePath,
                $EnableLogging,
                $EnableEmailNotification,
                $EmailSettings,
                $File,
                $EnableHeartbeat,
                $EnableProcessedFilesLog,
                $ProcessedFilesLog,
                $LogQueue,
                $LogLevels,
                $ConsoleLogLevel
            )

            # Clone EmailSettings to avoid modifying the original object
            if ($EnableEmailNotification) {
                $ClonedEmailSettings = $EmailSettings.Clone()
            }

            # Logging Function within the job
            function Write-Log {
                param(
                    [string]$Message,
                    [string]$Level = "INFO"
                )
                $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $LogEntry = "[$Timestamp] [$Level] $Message"

                # Enqueue the log message
                $LogQueue.Enqueue(@($LogEntry, $Level))

                # Only write to console if message level is >= ConsoleLogLevel
                if ($LogLevels[$Level] -ge $LogLevels[$ConsoleLogLevel]) {
                    # Optionally, write to the console with different colors based on level
                    switch ($Level) {
                        "INFO"    { Write-Host $LogEntry -ForegroundColor Green }
                        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
                        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
                        "DEBUG"   { Write-Host $LogEntry -ForegroundColor Cyan }
                        default   { Write-Host $LogEntry }
                    }
                }
            }

            # Email Notification Function
            function Send-ErrorEmail {
                param(
                    [string]$ErrorMessage
                )
                if ($EnableEmailNotification) {
                    $ClonedEmailSettings.Subject = "$($ClonedEmailSettings.Subject): Error Occurred"
                    $ClonedEmailSettings.Body    = $ErrorMessage
                    try {
                        Send-MailMessage @ClonedEmailSettings
                        Write-Log "Error email sent successfully."
                    } catch {
                        Write-Log "Failed to send error email. Error: $_" "ERROR"
                    }
                }
            }

            # Heartbeat Update Function
            function Update-Heartbeat {
                if ($EnableHeartbeat) {
                    try {
                        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        Set-Content -Path $HeartbeatFile -Value $Timestamp
                        Write-Log "Heartbeat updated." "DEBUG"
                    } catch {
                        Write-Log "Failed to update heartbeat: $_" "ERROR"
                    }
                }
            }

            # Processed Files Logging Function
            function Log-ProcessedFile {
                param(
                    [string]$File,
                    [string]$Status = "success"  # New parameter with default value
                )
                if ($EnableProcessedFilesLog) {
                    try {
                        # Initialize processedfiles.log with headers if it doesn't exist
                        if (-not (Test-Path -Path $ProcessedFilesLog)) {
                            "Timestamp,File,Status" | Out-File -FilePath $ProcessedFilesLog -Encoding utf8
                        }
                        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        $LogEntry = "$Timestamp,$File,$Status"  # Include status in log entry
                        # Write directly to processedfiles.log
                        Add-Content -Path $ProcessedFilesLog -Value $LogEntry
                        Write-Log "Logged processed file: $FilePath ($Status)" "DEBUG"
                    } catch {
                        Write-Log "Failed to write to processedfiles.log: $_" "ERROR"
                    }
                }
            }

            # Start processing
            Write-Log "Worker job started for file: $FilePath" "INFO"

            # Wait until the file is fully accessible
            while ($true) {
                try {
                    Write-Log "Attempting to open file: $FilePath" "DEBUG"
                    $Stream = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
                    $Stream.Close()
                    Write-Log "File is accessible: $FilePath" "DEBUG"
                    break
                } catch {
                    Write-Log "File is locked or inaccessible: $FilePath. Retrying in 500ms." "WARNING"
                    Start-Sleep -Milliseconds 500
                }
            }

            # Properly quote the file path
            $QuotedFilePath = '"' + $FilePath + '"'
            $ScanmailCommand = "-c 'load SMDdbs -manual $QuotedFilePath'"

            try {
                Write-Log "Executing ScanMail command: $ScanmailCommand" "DEBUG"
                # Execute the scan command
                & $ConsolePath $ScanmailCommand
                Write-Log "Scanned NSF file: $FilePath" "INFO"
                # Log the processed file with 'success' status
                Log-ProcessedFile -File $FilePath -Status "success"
            } catch {
                $ErrorMessage = "Failed to scan $FilePath. Error: $_"
                Write-Log $ErrorMessage "ERROR"
                # Log the processed file with 'failed' status
                Log-ProcessedFile -File $FilePath -Status "failed"
                Send-ErrorEmail $ErrorMessage
            }

            # Optionally, update heartbeat after processing
            Update-Heartbeat

            Write-Log "Worker job completed for file: $FilePath" "INFO"
        } -ArgumentList `
            $ConsolePath, `
            $FilePath, `
            $EnableLogging, `
            $EnableEmailNotification, `
            $EmailSettings, `
            $HeartbeatFile, `
            $EnableHeartbeat, `
            $EnableProcessedFilesLog, `
            $ProcessedFilesLog, `
            $Global:LogQueue, `
            $Global:LogLevels, `
            $ConsoleLogLevel `
            -Name "WorkerJob_$(Get-Random -Maximum 10000)"

        # Add the job to the active jobs array
        $Global:ActiveWorkerJobs += $Job

        global:Write-Log "Started Worker Job for file: $FilePath with Job ID: $($Job.Id)" "INFO"
    }

    # Start the LogProcessorJob
    $Global:LogProcessorJob = Start-ThreadJob -ScriptBlock {
        param(
            $LogQueue,
            $LogFilePath,
            $EnableLogging,
            $FileLogLevel,
            $LogLevels
        )

        while ($true) {
            $LogEntryPair = $null
            if ($LogQueue.TryDequeue([ref]$LogEntryPair)) {
                $LogEntry = $LogEntryPair[0]
                $Level = $LogEntryPair[1]

                if ($EnableLogging -and $LogFilePath -and ($LogLevels[$Level] -ge $LogLevels[$FileLogLevel])) {
                    try {
                        Add-Content -Path $LogFilePath -Value $LogEntry
                    } catch {
                        Write-Host "Failed to write log entry: $_" -ForegroundColor Red
                    }
                }
            }
            Start-Sleep -Milliseconds 100
        }
    } -ArgumentList `
        $Global:LogQueue, `
        $LogFilePath, `
        $EnableLogging, `
        $FileLogLevel, `
        $Global:LogLevels `
        -Name "LogProcessorJob"

    # Wait for LogProcessorJob to start
    do {
        Start-Sleep -Milliseconds 500
        $LogProcessorJob = Get-Job -Name "LogProcessorJob"
    } while ($LogProcessorJob.State -eq 'NotStarted' -or $LogProcessorJob.State -eq 'Disconnected')

    if ($LogProcessorJob.State -eq 'Running') {
        global:Write-Log "Log processor job started successfully." "INFO"
    } else {
        global:Write-Log "Log processor job failed to start. State: $($LogProcessorJob.State)" "ERROR"
    }

    # Set up the FileSystemWatcher
    $Watcher = New-Object System.IO.FileSystemWatcher
    $Watcher.Path = $Folder
    $Watcher.Filter = "*.nsf"
    $Watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName
    $Watcher.IncludeSubdirectories = $true
    $Watcher.InternalBufferSize = 65536  # Increase buffer size to handle more events
    $Watcher.EnableRaisingEvents = $true

    $FileCreatedEvent = Register-ObjectEvent $Watcher Created -SourceIdentifier FileCreatedEvent -Action {
        try {
            # Access the full path using the automatic $Event variable
            $FilePath = $Event.SourceEventArgs.FullPath

            # Log that the event was triggered
            global:Write-Log "FileSystemWatcher triggered for file: $FilePath" "INFO"

            # Get the current time
            $CurrentTime = Get-Date

            # Log the current time
            global:Write-Log "Current Time: $CurrentTime" "DEBUG"

            # Define a timeframe to consider events as duplicates (e.g., 5 seconds)
            $DuplicateTimeframe = [TimeSpan]::FromSeconds(5)
            global:Write-Log "Duplicate Timeframe: $DuplicateTimeframe" "DEBUG"

            # Identify keys to remove (files enqueued more than $DuplicateTimeframe ago)
            $KeysToRemove = $Global:ProcessedFiles.Keys | Where-Object {
                $Global:ProcessedFiles[$_] -le $CurrentTime.Add(-$DuplicateTimeframe)
            }

            # Log the keys identified for removal
            if ($KeysToRemove.Count -gt 0) {
                global:Write-Log "Keys to Remove: $($KeysToRemove -join ', ')" "DEBUG"
            } else {
                global:Write-Log "No keys to remove based on the timeframe." "DEBUG"
            }

            # Remove old entries
            foreach ($Key in $KeysToRemove) {
                $Global:ProcessedFiles.Remove($Key) | Out-Null
                global:Write-Log "Removed old entry from ProcessedFiles: $Key" "DEBUG"
            }

            # Check if the file was recently processed
            if (-not $Global:ProcessedFiles.ContainsKey($FilePath)) {
                global:Write-Log "Enqueued new NSF file: $FilePath" "INFO"
                $Global:Queue.Enqueue($FilePath)
                $Global:ProcessedFiles[$FilePath] = $CurrentTime
                global:Write-Log "Added $FilePath to ProcessedFiles with timestamp $CurrentTime" "DEBUG"
            } else {
                global:Write-Log "Duplicate event detected for file: $FilePath. Ignoring." "DEBUG"
            }
        } catch {
            global:Write-Log "Error in FileCreatedEvent handler: $_" "ERROR"
        }
    }

    # =====================================
    #         Heartbeat Mechanism
    # =====================================

    if ($EnableHeartbeat) {
        global:Write-Log "Heartbeat mechanism enabled." "INFO"

        # Set up the heartbeat ThreadJob with correct parameters
        $Global:HeartbeatJob = Start-ThreadJob -ScriptBlock {
            param(
                $HeartbeatInterval,
                $EnableHeartbeat,
                $HeartbeatFile,
                $LogQueue,
                $LogLevels,
                $ConsoleLogLevel
            )
            # Define the Write-Log function within the ThreadJob
            function Write-Log {
                param(
                    [string]$Message,
                    [string]$Level = "INFO"
                )
                $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $LogEntry = "[$Timestamp] [$Level] $Message"

                # Enqueue the log message
                $LogQueue.Enqueue(@($LogEntry, $Level))

                # Only write to console if message level is >= ConsoleLogLevel
                if ($LogLevels[$Level] -ge $LogLevels[$ConsoleLogLevel]) {
                    # Optionally, write to the console with different colors based on level
                    switch ($Level) {
                        "INFO"    { Write-Host $LogEntry -ForegroundColor Green }
                        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
                        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
                        "DEBUG"   { Write-Host $LogEntry -ForegroundColor Cyan }
                        default   { Write-Host $LogEntry }
                    }
                }
            }

            # Heartbeat Update Function within the ThreadJob
            function Update-Heartbeat {
                if ($EnableHeartbeat) {
                    try {
                        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        Set-Content -Path $HeartbeatFile -Value $Timestamp
                        Write-Log "Heartbeat updated." "DEBUG"
                    } catch {
                        Write-Log "Failed to update heartbeat: $_" "ERROR"
                    }
                }
            }

            # Initialize heartbeat immediately
            Update-Heartbeat

            # Start the heartbeat loop
            while ($EnableHeartbeat) {
                Start-Sleep -Milliseconds $HeartbeatInterval
                try {
                    Update-Heartbeat
                } catch {
                    Write-Log "Heartbeat ThreadJob encountered an error: $_" "ERROR"
                }
            }
        } -ArgumentList `
            $HeartbeatInterval, `
            $EnableHeartbeat, `
            $HeartbeatFile, `
            $Global:LogQueue, `
            $Global:LogLevels, `
            $ConsoleLogLevel `
            -Name "HeartbeatJob"

        global:Write-Log "Heartbeat ThreadJob started." "INFO"
    }

    # =====================================
    #         Self-Check Mechanism
    # =====================================

    if ($EnableSelfCheck) {
        global:Write-Log "Self-check mechanism enabled." "INFO"

        # Set up the self-check ThreadJob with correct parameters
        $Global:SelfCheckJob = Start-ThreadJob -ScriptBlock {
            param(
                $SelfCheckInterval,
                $MaxConcurrentScans,
                $ActiveWorkerJobs,
                $LogQueue,
                $LogLevels,
                $ConsoleLogLevel
            )

            # Define the Write-Log function within the ThreadJob
            function Write-Log {
                param(
                    [string]$Message,
                    [string]$Level = "INFO"
                )
                $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $LogEntry = "[$Timestamp] [$Level] $Message"

                # Enqueue the log message
                $LogQueue.Enqueue(@($LogEntry, $Level))

                # Only write to console if message level is >= ConsoleLogLevel
                if ($LogLevels[$Level] -ge $LogLevels[$ConsoleLogLevel]) {
                    # Optionally, write to the console with different colors based on level
                    switch ($Level) {
                        "INFO"    { Write-Host $LogEntry -ForegroundColor Green }
                        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
                        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
                        "DEBUG"   { Write-Host $LogEntry -ForegroundColor Cyan }
                        default   { Write-Host $LogEntry }
                    }
                }
            }

            # Start the self-check loop
            while ($EnableHeartbeat) {
                Start-Sleep -Milliseconds $SelfCheckInterval
                try {
                    # Log the self-check initiation
                    Write-Log "Running self-check..." "INFO"

                    # Get all active worker jobs that are not running
                    $WorkerJobs = Get-Job | Where-Object { $_.Name -like 'WorkerJob_*' -and $_.State -ne 'Running' }

                    foreach ($Job in $WorkerJobs) {
                        $WarningMessage = "Job ID $($Job.Id) ($($Job.Name)) is not running. Removing..."
                        Write-Log $WarningMessage "WARNING"
                        Remove-Job $Job -Force -ErrorAction SilentlyContinue
                        # Optionally, restart the job if needed
                    }
                } catch {
                    Write-Log "Self-check mechanism encountered an error: $_" "ERROR"
                }
            }
        } -ArgumentList `
            $SelfCheckInterval, `
            $MaxConcurrentScans, `
            $Global:ActiveWorkerJobs, `
            $Global:LogQueue, `
            $Global:LogLevels, `
            $ConsoleLogLevel `
            -Name "SelfCheckJob"

        global:Write-Log "Self-check ThreadJob started." "INFO"
    }

    # =====================================
    #         Script Execution Start
    # =====================================

    try {
        global:Write-Log "Monitoring directory: $Folder" "INFO"
        global:Write-Log "Script started successfully." "INFO"

        # Main loop to assign files to worker jobs
        while ($true) {
            # Check if there are files in the queue and if we can start new worker jobs
            while (($Global:Queue.Count -gt 0) -and ($Global:ActiveWorkerJobs.Count -lt $MaxConcurrentScans)) {
                # Dequeue a file
                $FilePath = $null
                if ($Global:Queue.TryDequeue([ref]$FilePath)) {
                    # Start a worker job for the file
                    Start-WorkerJob -FilePath $FilePath
                }
            }

            # Clean up completed worker jobs
            $Global:ActiveWorkerJobs = $Global:ActiveWorkerJobs | Where-Object { $_.State -eq 'Running' }

            # Sleep for a short interval before checking again
            Start-Sleep -Seconds 1
        }
    } finally {
        # Ensure cleanup is called even if the loop is exited unexpectedly
        global:Cleanup-Script
    }

} catch {
    # Handle any unforeseen errors
    $ErrorMessage = "An unexpected error occurred: $_"
    global:Write-Log $ErrorMessage "ERROR"
    global:Send-ErrorEmail $ErrorMessage
    global:Cleanup-Script
    exit 1
}