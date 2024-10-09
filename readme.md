# Installation and Configuration Guide

This guide provides step-by-step instructions for configuring and installing the **Trend Micro ScanMail for Domino Instant NSF Scanner Script** as a Windows Task Scheduler task that runs in the background upon system boot.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Script Configuration](#script-configuration)
   - [Universal Settings](#universal-settings)
   - [Configuration Toggles and Features](#configuration-toggles-and-features)
3. [Installing the Script](#installing-the-script)
   - [Step 1: Place the Script File](#step-1-place-the-script-file)
   - [Step 2: Configure Execution Policy](#step-2-configure-execution-policy)
   - [Step 3: Create a Scheduled Task](#step-3-create-a-scheduled-task)
4. [Verification](#verification)
5. [Troubleshooting](#troubleshooting)
6. [Security Considerations](#security-considerations)

---

## Prerequisites

Before installing and configuring the script, ensure the following prerequisites are met:

- **Operating System**: Windows Server or Windows Desktop OS with PowerShell installed (version 5.0 or higher is recommended).
- **Permissions**: Administrator privileges to create scheduled tasks and modify system configurations.
- **Trend Micro ScanMail for Domino**: Ensure it is installed and configured correctly on the system.
- **SMTP Server**: An SMTP server is available for sending email notifications if email alerts are enabled.

---

## Script Configuration

### Universal Settings

These settings are essential for the script to function correctly. Update the placeholders with the actual paths and desired values.

```powershell
# Directory to monitor for new NSF files.
$Folder = "C:\path\to\your\NSF\directory"

# Path to the ScanMail console executable.
$ScanmailPath = "C:\path\to\domino\console.exe"

# Maximum number of concurrent scans to run simultaneously.
$MaxConcurrentScans = 4
```

**Instructions:**

- **$Folder**: Replace `"C:\path\to\your\NSF\directory"` with the path to the directory where NSF files are placed for scanning.
- **$ScanmailPath**: Replace `"C:\path\to\domino\console.exe"` with the actual path to the `console.exe` file for ScanMail.
- **$MaxConcurrentScans**: Set this to the number of concurrent scans you want to allow. This depends on your system's resources.

### Configuration Toggles and Features

The script includes several optional features that can be enabled or disabled. Each feature has associated configuration settings.

#### 1. Logging

- **Toggle**: `$EnableLogging = $true`
- **Description**: Enables logging of script activity to a file for auditing and troubleshooting.

**Configuration:**

```powershell
if ($EnableLogging) {
    # Path to the log file where script activity will be recorded.
    $LogFilePath = "C:\path\to\logfile.log"
}
```

- **$LogFilePath**: Set the path where you want the log file to be stored.

#### 2. Email Notifications on Errors

- **Toggle**: `$EnableEmailNotification = $true`
- **Description**: Sends an email notification whenever an error occurs during processing.

**Configuration:**

```powershell
if ($EnableEmailNotification) {
    $EmailSettings = @{
        To         = "your.email@domain.com"          # Recipient email address.
        From       = "script.alerts@domain.com"       # Sender email address.
        Subject    = "ScanMail Script Error Notification"  # Email subject line.
        SmtpServer = "smtp.yourdomain.com"            # SMTP server address.
        # Uncomment and set the following lines if authentication is required
        # Credential = Get-Credential                 # Credentials for SMTP server authentication.
        # UseSsl     = $true                          # Enable SSL for SMTP connection.
    }
}
```

- **$EmailSettings.To**: Set to the email address that should receive error notifications.
- **$EmailSettings.From**: Set to the email address that will appear as the sender.
- **$EmailSettings.SmtpServer**: Set to your SMTP server's address.
- **Authentication**: If your SMTP server requires authentication, uncomment and configure the `Credential` and `UseSsl` lines.

#### 3. Heartbeat Mechanism

- **Toggle**: `$EnableHeartbeat = $true`
- **Description**: Updates a heartbeat file periodically to indicate that the script is running.

**Configuration:**

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

#### 4. Self-Check Mechanism

- **Toggle**: `$EnableSelfCheck = $true`
- **Description**: Periodically checks worker jobs and restarts them if they have stopped running.

**Configuration:**

```powershell
if ($EnableSelfCheck) {
    # Interval in milliseconds for checking worker jobs (e.g., 300000ms = 5 minutes).
    $SelfCheckInterval = 300000
}
```

- **$SelfCheckInterval**: Adjust the interval for how often the script checks and restarts jobs.

#### 5. Processed Files Logging

- **Toggle**: `$EnableProcessedFilesLog = $true`
- **Description**: Logs each successfully processed file to a log file for record-keeping.

**Configuration:**

```powershell
if ($EnableProcessedFilesLog) {
    # Path to the log file where processed files are recorded.
    $ProcessedFilesLog = "C:\path\to\processedfiles.log"
}
```

- **$ProcessedFilesLog**: Set the path for the processed files log.

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

     ```plaintext
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

---

## Verification

After setting up the task, verify that the script is running correctly.

1. **Reboot the System**:

   - Restart the computer to trigger the task at startup.

2. **Check Task Scheduler**:

   - Open Task Scheduler.
   - Locate your task and check the **Last Run Time** and **Last Run Result**.
   - A **Last Run Result** of `0x0` indicates success.

3. **Verify Logs**:

   - If logging is enabled, check the log file specified in `$LogFilePath`.
   - Confirm that the script has started and is monitoring the specified directory.

4. **Test File Detection**:

   - Place a test NSF file in the monitored directory.
   - Observe if the file is processed and logged accordingly.

5. **Check Heartbeat File**:

   - If the heartbeat mechanism is enabled, verify that the heartbeat file is being updated at the specified intervals.

---

## Troubleshooting

- **Script Not Running**:

  - Ensure the task is enabled and scheduled correctly.
  - Check the execution policy settings.
  - Verify that the user account running the task has the necessary permissions.

- **Errors in Logs**:

  - Review the log file for any error messages.
  - Ensure all file paths and configurations are correct.

- **Email Notifications Not Sent**:

  - Verify SMTP server settings.
  - Check if the SMTP server requires authentication and configure accordingly.
  - Ensure that network connectivity to the SMTP server is available.

- **Heartbeat File Not Updating**:

  - Confirm that the heartbeat mechanism is enabled.
  - Check the specified path for the heartbeat file.

- **Processed Files Not Logged**:

  - Ensure the processed files logging is enabled.
  - Verify the path to the processed files log.

---

## Security Considerations

- **Execution Policy**: Adjusting the execution policy can have security implications. Set it to the most restrictive setting that still allows the script to run.

- **Credentials**:

  - If using credentials for SMTP authentication, consider using secure methods to store and retrieve credentials, such as using the `Get-Credential` cmdlet and storing credentials securely.

- **File Permissions**:

  - Ensure that the script file, log files, and any directories accessed by the script have appropriate permissions to prevent unauthorized access.

- **Script Integrity**:

  - Keep the script in a secure location and limit write access to prevent unauthorized modifications.

---

**Note**: This guide is intended to assist in the deployment of the script within an organizational environment. Always follow your organization's IT policies and guidelines when configuring scripts and scheduled tasks.