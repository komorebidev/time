# ============================================================
# Interactive Scheduled Task Setup (Idle-Triggered)
#
# Creates a Windows Scheduled Task for a PowerShell script.
#
# Features:
#   - Interactive script-path prompt
#   - Interactive task-name prompt
#   - Interactive idle-duration prompt (defaults to 20 minutes)
#   - Runs PowerShell hidden
#   - Runs only when the current user is logged in
#   - Prevents overlapping executions
#   - Runs ONLY when the computer is idle (respects user activity)
#   - Restarts after failures
#   - Can immediately test the task
# ============================================================

$ErrorActionPreference = "Stop"


# ============================================================
# INTRO
# ============================================================

Clear-Host

Write-Host ""
Write-Host "=============================================="
Write-Host " PowerShell Scheduled Task Setup (Idle-Based)"
Write-Host "=============================================="
Write-Host ""

Write-Host "This will create a scheduled task that runs"
Write-Host "a PowerShell script silently when you are idle."
Write-Host ""

Write-Host "The task will:"
Write-Host "  - Run only when you are logged in"
Write-Host "  - Run with no visible PowerShell window"
Write-Host "  - Prevent overlapping instances"
Write-Host "  - Run ONLY when the PC has been idle for the set duration"
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
# ASK FOR IDLE DURATION
# ============================================================

Write-Host ""

$defaultIdleMinutes = 20

while ($true) {

    $idleInput = Read-Host `
        "Run after how many minutes of computer idle time? [$defaultIdleMinutes]"


    # Blank = use default

    if ([string]::IsNullOrWhiteSpace($idleInput)) {

        $idleMinutes = $defaultIdleMinutes

        break
    }


    # Safely convert the user's input to an integer

    try {

        $idleMinutes = [int]$idleInput

    }
    catch {

        $idleMinutes = 0
    }


    if ($idleMinutes -ge 1) {

        break
    }


    Write-Host ""
    Write-Host "ERROR: Please enter a whole number of at least 1."
}


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

Write-Host "Condition:"
Write-Host "  Runs after $idleMinutes minute(s) of system idle time"

Write-Host ""

Write-Host "Windows user:"
Write-Host "  $env:USERDOMAIN\$env:USERNAME"

Write-Host ""

Write-Host "Run mode:"
Write-Host "  Only when user is logged in"

Write-Host ""

Write-Host "PowerShell window:"
Write-Host "  Hidden"

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


    if (
        $replace -match "^(Y|y|Yes|yes)$"
    ) {

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
# ============================================================

$powerShellPath = Join-Path `
    $env:SystemRoot `
    "System32\WindowsPowerShell\v1.0\powershell.exe"


if (-not (Test-Path -LiteralPath $powerShellPath)) {

    throw "Windows PowerShell executable was not found."
}


# ============================================================
# CREATE ACTION
# ============================================================

$arguments = @(
    "-NoProfile"
    "-WindowStyle Hidden"
    "-ExecutionPolicy Bypass"
    "-File `"$scriptPath`""
) -join " "


$action = New-ScheduledTaskAction `
    -Execute $powerShellPath `
    -Argument $arguments


# ============================================================
# CREATE TRIGGER
#
# We set up a background recurrence matching your idle window 
# so Windows constantly checks if the criteria are met.
# ============================================================

$startTime = (Get-Date).AddMinutes(1)

$repetitionInterval = `
    New-TimeSpan -Minutes $idleMinutes

$repetitionDuration = `
    New-TimeSpan -Days 3650


$trigger = New-ScheduledTaskTrigger `
    -Once `
    -At $startTime `
    -RepetitionInterval $repetitionInterval `
    -RepetitionDuration $repetitionDuration


# ============================================================
# CREATE PRINCIPAL
# ============================================================

$currentUser = `
    "$env:USERDOMAIN\$env:USERNAME"


$principal = New-ScheduledTaskPrincipal `
    -UserId $currentUser `
    -LogonType Interactive `
    -RunLevel Limited


# ============================================================
# CREATE SETTINGS (Idle-Enforced)
#
# RunOnlyIfIdle
#   Forces Windows to skip execution unless system is idle.
# IdleDuration
#   Dictates how long the system must remain untouched before
#   the trigger criteria are validated.
# ============================================================

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -RunOnlyIfIdle `
    -IdleDuration (New-TimeSpan -Minutes $idleMinutes) `
    -MultipleInstances IgnoreNew `
    -RestartCount 3 `
    -RestartInterval (
        New-TimeSpan -Minutes 5
    )


# ============================================================
# REGISTER TASK
# ============================================================

Write-Host ""
Write-Host "Creating scheduled task..."

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description `
        "Automatically runs the Worklog Outlook ICS importer when idle."


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

Write-Host "Condition:"
Write-Host "  Runs after $idleMinutes minute(s) of system idle time"

Write-Host ""

Write-Host "PowerShell window:"
Write-Host "  Hidden"

Write-Host ""

Write-Host "First check starts:"
Write-Host "  Approximately $startTime"

Write-Host ""

Write-Host "Run mode:"
Write-Host "  Only while you are logged in and idle"

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


    Write-Host ""
    Write-Host "Task state:"
    Write-Host "  $($task.State)"

    Write-Host ""

    Write-Host "Task started successfully."

    Write-Host ""
    Write-Host "Because PowerShell is configured as hidden,"
    Write-Host "no PowerShell window should appear."
}


# ============================================================
# DONE
# ============================================================

Write-Host ""
Write-Host "Setup complete."