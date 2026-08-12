# ============================================================
# Interactive Scheduled Task Setup
#
# Creates a Windows Scheduled Task for a PowerShell script.
#
# Trigger:
#   Workstation unlock
#
# Features:
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
# Workstation unlock event:
#   Security Event ID 4801
#
# IMPORTANT:
#   Run this script from PowerShell as Administrator.
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

    # Remove surrounding quotes if pasted:
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

Write-Host "Event:"
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
# BUILD POWERSHELL COMMAND
# ============================================================

$arguments = @(
    "-NoProfile"
    "-WindowStyle Hidden"
    "-ExecutionPolicy Bypass"
    "-File `"$scriptPath`""
) -join " "


# ============================================================
# CREATE TASK SCHEDULER COM OBJECT
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

$taskDefinition = $scheduler.NewTask(0)


# ============================================================
# REGISTRATION INFORMATION
# ============================================================

$taskDefinition.RegistrationInfo.Description = `
    "Automatically runs the Worklog Outlook ICS importer whenever the workstation is unlocked."


# ============================================================
# PRINCIPAL
#
# InteractiveToken means the task runs using the currently
# logged-in user's interactive session.
#
# This is important for Outlook COM automation.
# ============================================================

$taskDefinition.Principal.UserId = $currentUser

$taskDefinition.Principal.LogonType = 3
# TASK_LOGON_INTERACTIVE_TOKEN

$taskDefinition.Principal.RunLevel = 0
# TASK_RUNLEVEL_LUA / Least Privilege


# ============================================================
# SETTINGS
#
# MultipleInstancesPolicy:
#   0 = Parallel
#   1 = Queue
#   2 = IgnoreNew
#
# We use 2 so a second unlock event doesn't start another
# copy while the previous execution is still running.
# ============================================================

$taskDefinition.Settings.MultipleInstances = 2


# ============================================================
# BATTERY SETTINGS
# ============================================================

$taskDefinition.Settings.DisallowStartIfOnBatteries = $false

$taskDefinition.Settings.StopIfGoingOnBatteries = $false


# ============================================================
# START WHEN AVAILABLE
# ============================================================

$taskDefinition.Settings.StartWhenAvailable = $true


# ============================================================
# EXECUTION TIME LIMIT
#
# PT0S means no execution time limit.
# ============================================================

$taskDefinition.Settings.ExecutionTimeLimit = "PT0S"


# ============================================================
# RESTART AFTER FAILURE
#
# Restart up to 3 times.
# Wait 5 minutes between attempts.
# ============================================================

$taskDefinition.Settings.RestartCount = 3

$taskDefinition.Settings.RestartInterval = "PT5M"


# ============================================================
# CREATE EVENT TRIGGER
#
# Windows Security Event 4801:
#
#   "The workstation was unlocked."
#
# The trigger watches the Security event log.
# ============================================================

$trigger = $taskDefinition.Triggers.Create(0)

# 0 = TASK_TRIGGER_EVENT


$trigger.Enabled = $true


# ============================================================
# EVENT SUBSCRIPTION
#
# This XML selects:
#
#   Log: Security
#   Provider: Microsoft-Windows-Security-Auditing
#   Event ID: 4801
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
# Arguments:
#   -NoProfile
#   -WindowStyle Hidden
#   -ExecutionPolicy Bypass
#   -File "script.ps1"
# ============================================================

$action = $taskDefinition.Actions.Create(0)

# 0 = TASK_ACTION_EXEC


$action.Path = $powerShellPath

$action.Arguments = $arguments


# ============================================================
# REGISTER TASK
# ============================================================

Write-Host ""
Write-Host "Registering scheduled task..."

# TASK_CREATE_OR_UPDATE = 6
#
# TASK_LOGON_INTERACTIVE_TOKEN = 3

$TASK_CREATE_OR_UPDATE = 6

$TASK_LOGON_INTERACTIVE_TOKEN = 3


try {

    $rootFolder.RegisterTaskDefinition(
        $taskName,
        $taskDefinition,
        $TASK_CREATE_OR_UPDATE,
        $currentUser,
        $null,
        $TASK_LOGON_INTERACTIVE_TOKEN,
        $null
    )

}
catch {

    Write-Host ""
    Write-Host "=============================================="
    Write-Host " ERROR"
    Write-Host "=============================================="
    Write-Host ""

    Write-Host "Task registration failed."
    Write-Host ""

    Write-Host $_.Exception.Message

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

    Write-Host "Task started successfully."

    Write-Host ""
    Write-Host "Because PowerShell is configured as hidden,"
    Write-Host "no PowerShell window should appear."

    Write-Host ""
    Write-Host "You can also lock and unlock Windows"
    Write-Host "to test the Event ID 4801 trigger."
}


# ============================================================
# DONE
# ============================================================

Write-Host ""
Write-Host "Setup complete."
Write-Host ""

Write-Host "The script will now run whenever"
Write-Host "the workstation is unlocked."