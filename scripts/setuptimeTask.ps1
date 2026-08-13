# ============================================================
# Interactive Scheduled Task Setup (Idle/Interval + Lock Check)
#
# Creates a Windows Scheduled Task for a PowerShell script.
#
# Features:
#   - Interactive script-path prompt
#   - Interactive task-name prompt
#   - Interactive interval prompt (e.g., every 5 minutes)
#   - Interactive idle-duration prompt (e.g., idle check)
#   - Runs PowerShell hidden
#   - Runs only when the current user is logged in
#   - Prevents overlapping executions
#   - Skips / cancels if the PC is locked (via script guard clause)
#   - Restarts after failures
# ============================================================

$ErrorActionPreference = "Stop"


# ============================================================
# INTRO
# ============================================================

Clear-Host

Write-Host ""
Write-Host "=============================================="
Write-Host " PowerShell Scheduled Task Setup (Interval + Idle)"
Write-Host "=============================================="
Write-Host ""

Write-Host "This will create a scheduled task that checks your"
Write-Host "script periodically, ensuring it only runs when you"
Write-Host "are logged in and your PC is not locked."
Write-Host ""


# ============================================================
# ASK FOR SCRIPT PATH
# ============================================================

while ($true) {

    Write-Host ""

    $scriptPath = Read-Host `
        "Enter the full path to the PowerShell script to schedule"

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
# ASK FOR INTERVAL (e.g., Every X minutes)
# ============================================================

Write-Host ""

$defaultIntervalMinutes = 5

while ($true) {

    $intervalInput = Read-Host `
        "Run every how many minutes? [$defaultIntervalMinutes]"

    if ([string]::IsNullOrWhiteSpace($intervalInput)) {

        $intervalMinutes = $defaultIntervalMinutes

        break
    }

    try {

        $intervalMinutes = [int]$intervalInput

    }
    catch {

        $intervalMinutes = 0
    }

    if ($intervalMinutes -ge 1) {

        break
    }

    Write-Host ""
    Write-Host "ERROR: Please enter a whole number of at least 1."
}


# ============================================================
# ASK FOR IDLE DURATION (Optional idle requirement)
# ============================================================

Write-Host ""

$defaultIdleMinutes = 0

while ($true) {

    $idleInput = Read-Host `
        "Require how many minutes of system idle time? (Enter 0 for none) [$defaultIdleMinutes]"

    if ([string]::IsNullOrWhiteSpace($idleInput)) {

        $idleMinutes = $defaultIdleMinutes

        break
    }

    try {

        $idleMinutes = [int]$idleInput

    }
    catch {

        $idleMinutes = -1
    }

    if ($idleMinutes -ge 0) {

        break
    }

    Write-Host ""
    Write-Host "ERROR: Please enter 0 or a positive whole number."
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

Write-Host "Schedule:"
Write-Host "  Triggers every $intervalMinutes minute(s)"

if ($idleMinutes -gt 0) {

    Write-Host "Idle Requirement:"
    Write-Host "  Requires $idleMinutes minute(s) of system idle time"
}

Write-Host ""

Write-Host "Windows user:"
Write-Host "  $env:USERDOMAIN\$env:USERNAME"

Write-Host ""

Write-Host "Run mode:"
Write-Host "  Only when user is logged in (Exits if locked)"

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
# CREATE TRIGGER (Repeating Interval)
# ============================================================

$startTime = (Get-Date).AddMinutes(1)

$trigger = New-ScheduledTaskTrigger `
    -Once `
    -At $startTime `
    -RepetitionInterval (New-TimeSpan -Minutes $intervalMinutes) `
    -RepetitionDuration (New-TimeSpan -Days 3650)


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
# CREATE SETTINGS
# ============================================================

if ($idleMinutes -gt 0) {

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -RunOnlyIfIdle `
        -IdleDuration (New-TimeSpan -Minutes $idleMinutes) `
        -MultipleInstances IgnoreNew `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 5)
}
else {

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -MultipleInstances IgnoreNew `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 5)
}


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
        "Automatically runs the Worklog Outlook ICS importer periodically, skipping if locked."


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

Write-Host "Schedule:"
Write-Host "  Triggers every $intervalMinutes minute(s)"

if ($idleMinutes -gt 0) {

    Write-Host "Idle Requirement:"
    Write-Host "  Requires $idleMinutes minute(s) of system idle time"
}

Write-Host ""

Write-Host "Lock Protection:"
Write-Host "  Script guard clause skips execution if PC is locked"

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
}


# ============================================================
# DONE
# ============================================================

Write-Host ""
Write-Host "Setup complete."