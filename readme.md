# Installation and Configuration Guide

This guide provides step-by-step instructions for configuring and installing the **Trend Micro ScanMail for Domino Instant NSF Scanner Script** as a Windows Task Scheduler task that runs in the background upon system boot.

## Features

This PowerShell script offers a wide range of features designed to deliver reliable, scalable, and efficient directory monitoring and file processing:

1. **Real-Time Directory Monitoring**: 
   - Continuously monitors a specified folder for new NSF database files.
   - Implements a queuing system to ensure all files are processed as they are added.

2. **Concurrency with Worker Jobs**: 
   - Supports multi-threaded file processing using PowerShell’s `Start-ThreadJob`, enabling parallel execution of tasks.
   - Includes job self-monitoring.

3. **Comprehensive Logging and Error Handling**: 
   - Logs system activity and errors at multiple severity levels (INFO, WARNING, ERROR).
   - Automatic email alerts for errors.

4. **Email Notifications**: 
   - Sends detailed email notifications for job completions, failures, or any critical events.

5. **Self-Check Mechanism**: 
   - A dedicated thread monitors job performance, restarting any failed or timed-out jobs to maintain reliability.

6. **Scalable Resource Management**: 
   - Configurable maximum number of concurrent scans (`$MaxConcurrentScans`) to optimize system performance based on resource availability.

7. **Heartbeat and Processed Files Tracking**: 
   - Tracks the health of the system using a heartbeat mechanism to allow external monitoring of script execution.
   - Logs processed files to prevent duplicate processing and to maintain a record of completed tasks.

8. **Cleanup Routines**: 
   - Ensures proper cleanup of system resources and temporary files upon script completion or exit.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Script Configuration](#script-configuration)
   - [Universal Settings](#universal-settings)
   - [Configuration Toggles and Features](#configuration-toggles-and-features)
3. [Installing the Script](#installing-the-script)
   - [Step 1: Place the Script File](#step-1-place-the-script-file)
   - [Step 2: Configure Execution Policy](#step-2-configure-execution-policy)
   - [Step 3: Create a Scheduled Task](#step-3-create-a-scheduled-task)
   - [Optional Step: Store SMTP Credentials](#optional-step-store-smtp-credentials)
4. [Verification](#verification)
5. [Troubleshooting](#troubleshooting)
6. [Security Considerations](#security-considerations)

---

## Prerequisites

Before installing and configuring the script, ensure the following prerequisites are met:

- **Operating System**: Windows Server or Windows Desktop OS with PowerShell installed (version 7.0 or higher is required).
  - **PowerShell Installation**: Use the following [`winget`](https://learn.microsoft.com/en-us/windows/package-manager/winget/) command to install the latest release version of PowerShell:
    ```
    winget install --id Microsoft.Powershell --source winget
    ```
- **Permissions**: Administrator privileges to create scheduled tasks and modify system configurations.
- **Trend Micro ScanMail for Domino**: Ensure it is installed and configured correctly on the system.
- **SMTP Server**: *(Optional)* An SMTP server is available for sending email notifications if email alerts are enabled.
- **PowerShell Module**: *(Optional)* The `CredentialManager` module must be installed if email notifications with authentication are enabled.

---

## Script Configuration

### Universal Settings

These settings are essential for the script to function correctly. Update the placeholders with the actual paths and desired values.

```powershell
# Directory to monitor for new NSF files.
$Folder = "C:\path\to\your\NSF\directory"

# Path to the Domino server console executable.
$ConsolePath = "C:\path\to\Domino\nserver.exe"

# Maximum number of concurrent scans to run simultaneously.
$MaxConcurrentScans = 4

# Console log level.
$ConsoleLogLevel = "INFO"  # Valid options: DEBUG, INFO, WARNING, ERROR
```

**Instructions:**

- **$Folder**: Replace `"C:\path\to\your\NSF\directory"` with the path to the directory where NSF files are placed for scanning.
- **$ConsolePath**: Replace `"C:\path\to\Domino\nserver.exe"` with the actual path to the Domino server console executable.
- **$MaxConcurrentScans**: Set this to the number of concurrent scans you want to allow. This depends on your system's resources.
- **$ConsoleLogLevel**: Set the desired console log level. Valid options are `DEBUG`, `INFO`, `WARNING`, and `ERROR`.

### Configuration Toggles and Features

The script includes several optional features that can be enabled or disabled. Each feature has associated configuration settings.

**By default, the following features are enabled:**

- **Self-Check Mechanism**

**Other features are disabled by default.**

#### 1. Logging

- **Toggle**: `$EnableLogging = $false` *(Disabled by default)*
- **Description**: Enables logging of script activity to a file for auditing and troubleshooting.

**Configuration (if enabled):**

```powershell
if ($EnableLogging) {
    # Path to the log file where script activity will be recorded.
    $LogFilePath = "C:\path\to\logfile.log"
    # File log level.
    $FileLogLevel = "INFO"  # Valid options: DEBUG, INFO, WARNING, ERROR
}
```

- **$LogFilePath**: Set the path where you want the log file to be stored.
- **$FileLogLevel**: Set the desired log level for the file output.

#### 2. Email Notifications on Errors

- **Toggle**: `$EnableEmailNotification = $false` *(Disabled by default)*
- **Description**: Sends an email notification whenever an error occurs during processing.

**Configuration (if enabled):**

```powershell
if ($EnableEmailNotification) {
    # Retrieve credentials from Credential Manager (if authentication is required)
    # $Credential = Get-StoredCredential -Target "SMTPCredential"

    $EmailSettings = @{
        To         = "your.email@domain.com"                 # Recipient email address.
        From       = "script.alerts@domain.com"              # Sender email address.
        Subject    = "ScanMail Script Error Notification"    # Email subject line.
        SmtpServer = "smtp.yourdomain.com"                   # SMTP server address.
        # Credential = $Credential                           # Uncomment if authentication is required.
        # UseSsl     = $true                                 # Enable SSL for SMTP connection if required.
        # Port       = 587                                   # SMTP port if different from default (25).
    }
}
```

**Instructions:**

- **EnableEmailNotification**: Set to `$true` to enable email notifications.
- **SMTP Configuration**:
  - **$EmailSettings.To**: Set to the email address that should receive error notifications.
  - **$EmailSettings.From**: Set to the email address that will appear as the sender.
  - **$EmailSettings.SmtpServer**: Set to your SMTP server's address.
- **Authentication**:
  - If your SMTP server requires authentication, uncomment and configure the `Credential` line.
  - **Note**: You need to store SMTP credentials (see [Optional Step: Store SMTP Credentials](#optional-step-store-smtp-credentials)) if email notifications are enabled and authentication is required.

#### 3. Heartbeat Mechanism

- **Toggle**: `$EnableHeartbeat = $false` *(Disabled by default)*
- **Description**: Updates a heartbeat file periodically to indicate that the script is running.

**Configuration (if enabled):**

```powershell
if ($EnableHeartbeat) {
    # Path to the heartbeat file.
    $HeartbeatFile = "C:\path\to\heartbeat.txt"
    # Interval in milliseconds for updating the heartbeat file (e.g., 60000ms = 60 seconds).
    $HeartbeatInterval = 60000
}
```

- **$HeartbeatFile**: Set the path where the heartbeat file will be created or updated.
- **$HeartbeatInterval**: Adjust the interval as needed.

#### 4. Processed Files Logging

- **Toggle**: `$EnableProcessedFilesLog = $false` *(Disabled by default)*
- **Description**: Logs each processed file to a log file for record-keeping.

**Configuration:**

```powershell
if ($EnableProcessedFilesLog) {
    # Path to the log file where processed files are recorded.
    $ProcessedFilesLog = "C:\path\to\processedfiles.log"
}
```

- **$ProcessedFilesLog**: Set the path for the processed files log.

#### 5. Self-Check Mechanism

- **Toggle**: `$EnableSelfCheck = $true` *(Enabled by default)*
- **Description**: Periodically checks worker jobs and restarts them if they have stopped running.

**Configuration:**

```powershell
if ($EnableSelfCheck) {
    # Interval in milliseconds for checking worker jobs (e.g., 300000ms = 5 minutes).
    $SelfCheckInterval = 300000
}
```

- **$SelfCheckInterval**: Adjust the interval for how often the script checks and restarts jobs.

---

## Installing the Script

### Step 1: Place the Script File

1. **Create a Directory for the Script**:

   Choose a directory to store the script file. For example:

   ```
   C:\Scripts\ScanMail\
   ```

2. **Save the Script**:

   - Copy the entire script into a text editor.
   - Save the file with a `.ps1` extension, e.g., `ScanMailMonitor.ps1`.
   - Ensure the file encoding is UTF-8 without BOM to avoid any execution issues.

### Step 2: Configure Execution Policy

To allow the script to run, you may need to adjust the PowerShell execution policy.

1. **Open PowerShell as Administrator**:

   - Right-click on the PowerShell icon and select **Run as Administrator**.

2. **Set Execution Policy**:

   ```powershell
   Set-ExecutionPolicy RemoteSigned
   ```

   - This allows scripts signed by a trusted publisher or scripts you write on the local computer to run.
   - **Note**: Be cautious when changing the execution policy. Consult your organization's security policies.

### Step 3: Create a Scheduled Task

Use the Windows Task Scheduler to run the script at system startup.

1. **Open Task Scheduler**:

   - Press `Win + R`, type `taskschd.msc`, and press `Enter`.

2. **Create a New Task**:

   - In the **Actions** pane, click **Create Task**.

3. **General Tab**:

   - **Name**: Enter a name for the task, e.g., `ScanMail NSF Monitor`.
   - **Description**: Provide a description if desired.
   - **Security Options**:
     - **Run whether user is logged on or not**: Select this option.
     - **Run with highest privileges**: Check this box.
     - **Configure for**: Select your Windows version.

4. **Triggers Tab**:

   - Click **New**.
   - **Begin the task**: Select **At startup**.
   - **Delay task for**: (Optional) Set a delay if needed.
   - Click **OK**.

5. **Actions Tab**:

   - Click **New**.
   - **Action**: Select **Start a program**.
   - **Program/script**: Enter `powershell.exe`.
   - **Add arguments (optional)**:

     ```
     -ExecutionPolicy Bypass -File "C:\Scripts\ScanMail\ScanMailMonitor.ps1"
     ```

     - Replace the path with the actual path to your script file.
   - Click **OK**.

6. **Conditions Tab**:

   - Adjust conditions as necessary. For example, you might uncheck **Start the task only if the computer is on AC power** if running on a laptop.

7. **Settings Tab**:

   - **Allow task to be run on demand**: Check this option.
   - **Stop the task if it runs longer than**: (Optional) Uncheck if you don't want the task to stop.
   - **If the task is already running, then the following rule applies**: Choose **Do not start a new instance**.

8. **OK and Save**:

   - Click **OK** to save the task.
   - If prompted, enter the credentials of the user account under which the task will run. This account should have the necessary permissions to execute the script and access required resources.

### Optional Step: Store SMTP Credentials

**This step is only necessary if you have enabled the Email Notifications on Errors feature and your SMTP server requires authentication.**

1. **Install the CredentialManager Module**:

   Open PowerShell as Administrator and run:

   ```powershell
   Install-Module -Name CredentialManager -Force
   ```

   - **Note**: You may need to adjust the execution policy temporarily to install the module:

     ```powershell
     Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
     ```

2. **Store SMTP Credentials in Credential Manager**:

   Run the following command in PowerShell (replace placeholders with your actual SMTP username and password):

   ```powershell
   New-StoredCredential -Target "SMTPCredential" -Username "your_smtp_username" -Password "your_smtp_password" -Type Generic
   ```

   - **"SMTPCredential"**: This is the name used to retrieve the credential in the script. Ensure it matches the name in the script's `Get-StoredCredential` line.
   - **"your_smtp_username"**: Replace with your SMTP username (e.g., `user@domain.com`).
   - **"your_smtp_password"**: Replace with your SMTP password.

   **Security Note**:

   - Ensure you run this command under the same user account that will execute the script.
   - The credentials are stored securely by Windows and are encrypted.

3. **Verify Credential Storage**:

   You can verify that the credentials are stored by running:

   ```powershell
   Get-StoredCredential -Target "SMTPCredential"
   ```

   - This should return the credential object without displaying the password in plain text.

---

## Verification

After setting up the task, verify that the script is running correctly.

1. **Reboot the System**:

   - Restart the computer to trigger the task at startup.

2. **Check Task Scheduler**:

   - Open Task Scheduler.
   - Locate your task and check the **Last Run Time** and **Last Run Result**.
   - A **Last Run Result** of `0x0` indicates success.

3. **Verify Functionality**:

   - **Self-Check Mechanism**: Since this feature is enabled by default, ensure that worker jobs are running correctly. Check for any entries in the event logs or console outputs if applicable.

4. **Optional Features Verification**:

   - **If you enabled any optional features (Logging, Email Notifications, Heartbeat Mechanism, Processed Files Logging), verify their functionality accordingly.**

   - **Logging**:
     - Check the log file specified in `$LogFilePath` for script activity.

   - **Email Notifications**:
     - Induce an error intentionally (e.g., temporarily set an incorrect path in `$ConsolePath`) to test if the script sends an email notification.
     - After testing, revert any intentional errors.

   - **Heartbeat File**:
     - If the heartbeat mechanism is enabled, verify that the heartbeat file is being updated at the specified intervals.

   - **Processed Files Log**:
     - If processed file logging is enabled, place a test NSF file in the monitored directory and check if it's logged properly in the `$ProcessedFilesLog`.

---

## Troubleshooting

- **Script Not Running**:

  - Ensure the task is enabled and scheduled correctly.
  - Check the execution policy settings.
  - Verify that the user account running the task has the necessary permissions.

- **Errors in Logs**:

  - Review the log file for any error messages.
  - Ensure all file paths and configurations are correct.
  - Check for typos in the script, especially in the paths and settings.

- **Email Notifications Not Sent**:

  - Verify that the **Email Notifications on Errors** feature is enabled (`$EnableEmailNotification = $true`).
  - Check the SMTP server settings in `$EmailSettings`.
  - If authentication is required, ensure SMTP credentials are stored correctly in Credential Manager and the `Credential` parameter is configured in the script.
  - Ensure that network connectivity to the SMTP server is available.
  - Check for any firewall or antivirus settings that might block SMTP traffic.

- **Heartbeat File Not Updating**:

  - Confirm that the heartbeat mechanism is enabled (`$EnableHeartbeat = $true`).
  - Check the specified path for the heartbeat file.
  - Verify that the script has write permissions to the location.

- **Processed Files Not Logged**:

  - Ensure the processed files logging is enabled (`$EnableProcessedFilesLog = $true`).
  - Verify the path to the processed files log.
  - Check for permission issues on the log file's directory.

---

## Security Considerations

- **Execution Policy**: Adjusting the execution policy can have security implications. Set it to the most restrictive setting that still allows the script to run.

- **Credentials**:

  - **Secure Storage**: If using SMTP authentication, store credentials securely using the Credential Manager.
  - **Access Control**: Ensure that only authorized users can access the credentials and that the script runs under the correct user account.

- **File Permissions**:

  - Ensure that the script file, log files, and any directories accessed by the script have appropriate permissions to prevent unauthorized access.

- **Script Integrity**:

  - Keep the script in a secure location and limit write access to prevent unauthorized modifications.

- **Credential Manager Module**:

  - The `CredentialManager` module should be installed from a trusted source. Use the PowerShell Gallery to ensure authenticity.

---

**Note**: This guide is intended to assist in the deployment of the script within an organizational environment. Always follow your organization's IT policies and guidelines when configuring scripts and scheduled tasks.