# ============================================================
# Interactive Scheduled Task Setup
#
# Creates a Windows Scheduled Task for a PowerShell script.
#
# TRIGGER:
#   Workstation unlock
#
# Workstation unlock event:
#   Security Event ID 4801
#
# FEATURES:
#   - Interactive script-path prompt
#   - Interactive task-name prompt
#   - Runs when workstation is unlocked
#   - Runs only when current user is logged in
#   - Runs PowerShell hidden
#   - Prevents overlapping executions
#   - Starts when available
#   - Restarts after failures
#   - Can immediately test the task
#
# IMPORTANT:
#   Run this setup script as Administrator.
# ============================================================

$ErrorActionPreference = "Stop"


# ============================================================
# INTRO
# ============================================================

Clear-Host

Write-Host ""
Write-Host "=============================================="
Write-Host " PowerShell Scheduled Task Setup"
Write-Host "=============================================="
Write-Host ""

Write-Host "This will create a scheduled task that runs"
Write-Host "a PowerShell script whenever the workstation"
Write-Host "is unlocked."
Write-Host ""

Write-Host "The task will:"
Write-Host "  - Run when the workstation is unlocked"
Write-Host "  - Run only when you are logged in"
Write-Host "  - Run with no visible PowerShell window"
Write-Host "  - Prevent overlapping instances"
Write-Host "  - Start when available"
Write-Host "  - Restart after failures"
Write-Host ""


# ============================================================
# ASK FOR SCRIPT PATH
# ============================================================

while ($true) {

    Write-Host ""

    $scriptPath = Read-Host `
        "Enter the full path to the PowerShell script to schedule"

    # Remove surrounding quotes if the user pasted:
    # "C:\Worklog\Import-WorklogICS.ps1"

    $scriptPath = $scriptPath.Trim().Trim('"')


    if ([string]::IsNullOrWhiteSpace($scriptPath)) {

        Write-Host ""
        Write-Host "ERROR: Please enter a script path."
        continue
    }


    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {

        Write-Host ""
        Write-Host "ERROR: File not found:"
        Write-Host "  $scriptPath"
        continue
    }


    $extension = [System.IO.Path]::GetExtension($scriptPath)


    if ($extension -ine ".ps1") {

        Write-Host ""
        Write-Host "ERROR: The file must be a .ps1 PowerShell script."
        continue
    }


    break
}


# ============================================================
# ASK FOR TASK NAME
# ============================================================

Write-Host ""

$defaultTaskName = "Worklog Outlook ICS Import"

$taskName = Read-Host `
    "Task name [$defaultTaskName]"


if ([string]::IsNullOrWhiteSpace($taskName)) {

    $taskName = $defaultTaskName
}


# ============================================================
# CURRENT USER
# ============================================================

$currentUser = "$env:USERDOMAIN\$env:USERNAME"


# ============================================================
# SHOW CONFIGURATION
# ============================================================

Write-Host ""
Write-Host "=============================================="
Write-Host " Configuration"
Write-Host "=============================================="
Write-Host ""

Write-Host "Task name:"
Write-Host "  $taskName"

Write-Host ""

Write-Host "Script:"
Write-Host "  $scriptPath"

Write-Host ""

Write-Host "Trigger:"
Write-Host "  Workstation unlock"

Write-Host ""

Write-Host "Windows event:"
Write-Host "  Security Event ID 4801"

Write-Host ""

Write-Host "Windows user:"
Write-Host "  $currentUser"

Write-Host ""

Write-Host "Run mode:"
Write-Host "  Only when user is logged in"

Write-Host ""

Write-Host "PowerShell window:"
Write-Host "  Hidden"

Write-Host ""

Write-Host "Multiple instances:"
Write-Host "  Ignore new instance"

Write-Host ""

Write-Host "Failure handling:"
Write-Host "  Restart up to 3 times"
Write-Host "  5 minutes between attempts"

Write-Host ""


# ============================================================
# CONFIRM
# ============================================================

$confirmation = Read-Host `
    "Create this scheduled task? [Y/n]"


if (
    $confirmation -and
    $confirmation -notmatch "^(Y|y|Yes|yes)$"
) {

    Write-Host ""
    Write-Host "Cancelled."

    exit 0
}


# ============================================================
# CHECK FOR EXISTING TASK
# ============================================================

$existingTask = Get-ScheduledTask `
    -TaskName $taskName `
    -ErrorAction SilentlyContinue


if ($existingTask) {

    Write-Host ""
    Write-Host "A scheduled task with this name already exists."

    $replace = Read-Host `
        "Replace the existing task? [y/N]"


    if ($replace -match "^(Y|y|Yes|yes)$") {

        Write-Host ""
        Write-Host "Removing existing task..."

        Unregister-ScheduledTask `
            -TaskName $taskName `
            -Confirm:$false

        Write-Host "Existing task removed."
    }
    else {

        Write-Host ""
        Write-Host "Cancelled."

        exit 0
    }
}


# ============================================================
# FIND WINDOWS POWERSHELL
#
# We intentionally use Windows PowerShell 5.1 because
# the target script uses Outlook COM automation.
# ============================================================

$powerShellPath = Join-Path `
    $env:SystemRoot `
    "System32\WindowsPowerShell\v1.0\powershell.exe"


if (-not (Test-Path -LiteralPath $powerShellPath)) {

    throw "Windows PowerShell executable was not found."
}


# ============================================================
# BUILD POWERSHELL ARGUMENTS
# ============================================================

$arguments = @(
    "-NoProfile"
    "-WindowStyle Hidden"
    "-ExecutionPolicy Bypass"
    "-File `"$scriptPath`""
) -join " "


# ============================================================
# CONNECT TO TASK SCHEDULER
# ============================================================

Write-Host ""
Write-Host "Connecting to Task Scheduler..."

$scheduler = New-Object -ComObject "Schedule.Service"

$scheduler.Connect()


# ============================================================
# GET ROOT TASK FOLDER
# ============================================================

$rootFolder = $scheduler.GetFolder("\")


# ============================================================
# CREATE TASK DEFINITION
# ============================================================

Write-Host "Creating task definition..."

$taskDefinition = $scheduler.NewTask(0)


# ============================================================
# REGISTRATION INFORMATION
# ============================================================

$taskDefinition.RegistrationInfo.Description = `
    "Automatically runs the Worklog Outlook ICS importer whenever the workstation is unlocked."


# ============================================================
# PRINCIPAL
#
# TASK_LOGON_INTERACTIVE_TOKEN = 3
#
# This means:
#
#   Run only when the user is logged in.
#
# This is important for Outlook COM automation.
# ============================================================

$TASK_LOGON_INTERACTIVE_TOKEN = 3

$TASK_RUNLEVEL_LUA = 0


$taskDefinition.Principal.UserId = $currentUser

$taskDefinition.Principal.LogonType = `
    $TASK_LOGON_INTERACTIVE_TOKEN

$taskDefinition.Principal.RunLevel = `
    $TASK_RUNLEVEL_LUA


# ============================================================
# SETTINGS
# ============================================================

# TASK_INSTANCES_IGNORE_NEW = 2

$TASK_INSTANCES_IGNORE_NEW = 2


$taskDefinition.Settings.MultipleInstances = `
    $TASK_INSTANCES_IGNORE_NEW


# Allow the task to run on battery.

$taskDefinition.Settings.DisallowStartIfOnBatteries = $false


# Don't stop it when switching to battery.

$taskDefinition.Settings.StopIfGoingOnBatteries = $false


# Run when the task becomes available.

$taskDefinition.Settings.StartWhenAvailable = $true


# No execution time limit.

$taskDefinition.Settings.ExecutionTimeLimit = "PT0S"


# ============================================================
# RESTART AFTER FAILURE
# ============================================================

$taskDefinition.Settings.RestartCount = 3

$taskDefinition.Settings.RestartInterval = "PT5M"


# ============================================================
# CREATE EVENT TRIGGER
#
# TASK_TRIGGER_EVENT = 0
#
# Event ID 4801:
#
#   The workstation was unlocked.
# ============================================================

Write-Host "Creating workstation unlock trigger..."

$trigger = $taskDefinition.Triggers.Create(0)

$trigger.Enabled = $true


# ============================================================
# EVENT QUERY
#
# Watch the Windows Security event log for:
#
# Provider:
#   Microsoft-Windows-Security-Auditing
#
# Event ID:
#   4801
# ============================================================

$trigger.Subscription = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[
        System[
          Provider[
            @Name='Microsoft-Windows-Security-Auditing'
          ]
          and
          (EventID=4801)
        ]
      ]
    </Select>
  </Query>
</QueryList>
"@


# ============================================================
# CREATE ACTION
#
# Execute:
#   Windows PowerShell 5.1
#
# Hidden:
#   -WindowStyle Hidden
# ============================================================

Write-Host "Creating PowerShell action..."

$action = $taskDefinition.Actions.Create(0)

# TASK_ACTION_EXEC = 0

$action.Path = $powerShellPath

$action.Arguments = $arguments


# ============================================================
# REGISTER TASK
#
# IMPORTANT:
#
# For TASK_LOGON_INTERACTIVE_TOKEN:
#
#   The user argument MUST be $null.
#
# The task's user is already specified through:
#
#   $taskDefinition.Principal.UserId
#
# This is the important fix for the previous
# "UnauthorizedAccessException" problem.
# ============================================================

Write-Host ""
Write-Host "Registering scheduled task..."


$TASK_CREATE_OR_UPDATE = 6


try {

    $registeredTask = $rootFolder.RegisterTaskDefinition(
        $taskName,
        $taskDefinition,
        $TASK_CREATE_OR_UPDATE,

        # IMPORTANT:
        # InteractiveToken requires NULL user/password here.

        $null,
        $null,

        $TASK_LOGON_INTERACTIVE_TOKEN,

        $null
    )

}
catch {

    Write-Host ""
    Write-Host "=============================================="
    Write-Host " TASK REGISTRATION FAILED"
    Write-Host "=============================================="
    Write-Host ""

    Write-Host "Error:"
    Write-Host $_.Exception.Message

    Write-Host ""

    Write-Host "HResult:"
    Write-Host $_.Exception.HResult

    Write-Host ""

    throw
}


# ============================================================
# SUCCESS
# ============================================================

Write-Host ""
Write-Host "=============================================="
Write-Host " SUCCESS"
Write-Host "=============================================="
Write-Host ""

Write-Host "Task:"
Write-Host "  $taskName"

Write-Host ""

Write-Host "Script:"
Write-Host "  $scriptPath"

Write-Host ""

Write-Host "Trigger:"
Write-Host "  Workstation unlock"

Write-Host ""

Write-Host "Event:"
Write-Host "  Security Event ID 4801"

Write-Host ""

Write-Host "PowerShell window:"
Write-Host "  Hidden"

Write-Host ""

Write-Host "Run mode:"
Write-Host "  Only while you are logged in"

Write-Host ""

Write-Host "Multiple instances:"
Write-Host "  Ignore new instance"

Write-Host ""

Write-Host "Failure handling:"
Write-Host "  Restart up to 3 times"
Write-Host "  5 minutes between attempts"

Write-Host ""


# ============================================================
# VERIFY TASK
# ============================================================

Write-Host "Verifying scheduled task..."

$verifyTask = Get-ScheduledTask `
    -TaskName $taskName `
    -ErrorAction Stop


Write-Host ""
Write-Host "Task state:"
Write-Host "  $($verifyTask.State)"

Write-Host ""


# ============================================================
# ASK WHETHER TO TEST NOW
# ============================================================

$runNow = Read-Host `
    "Run the scheduled task now as a test? [Y/n]"


if (
    [string]::IsNullOrWhiteSpace($runNow) -or
    $runNow -match "^(Y|y|Yes|yes)$"
) {

    Write-Host ""
    Write-Host "Starting task..."

    Start-ScheduledTask `
        -TaskName $taskName


    Start-Sleep -Seconds 3


    $task = Get-ScheduledTask `
        -TaskName $taskName


    $taskInfo = Get-ScheduledTaskInfo `
        -TaskName $taskName


    Write-Host ""
    Write-Host "Task state:"
    Write-Host "  $($task.State)"

    Write-Host ""

    Write-Host "Last run time:"
    Write-Host "  $($taskInfo.LastRunTime)"

    Write-Host ""

    Write-Host "Last run result:"
    Write-Host "  $($taskInfo.LastTaskResult)"

    Write-Host ""

    Write-Host "Manual test completed."

    Write-Host ""
    Write-Host "Now lock Windows with:"
    Write-Host ""
    Write-Host "  Win + L"
    Write-Host ""
    Write-Host "Then unlock the workstation."

    Write-Host ""
    Write-Host "The task should run automatically."
}


# ============================================================
# DONE
# ============================================================

Write-Host ""
Write-Host "=============================================="
Write-Host " Setup complete"
Write-Host "=============================================="
Write-Host ""

Write-Host "The task will run whenever"
Write-Host "the workstation is unlocked."

Write-Host ""

Write-Host "No interval is configured."
Write-Host ""
