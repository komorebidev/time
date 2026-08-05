# ============================================================
# Interactive Scheduled Task Setup
#
# Creates a Windows Task Scheduler task for a PowerShell script.
#
# Designed for Outlook COM automation:
#   - Runs only when current user is logged in
#   - Runs repeatedly at a chosen interval
#   - Prevents overlapping instances
#   - Restarts after failure
# ============================================================

$ErrorActionPreference = "Stop"


# ============================================================
# INTRO
# ============================================================

Write-Host ""
Write-Host "=============================================="
Write-Host " PowerShell Scheduled Task Setup"
Write-Host "=============================================="
Write-Host ""

Write-Host "This will create a scheduled task that runs"
Write-Host "a PowerShell script periodically."
Write-Host ""

Write-Host "The task will:"
Write-Host "  - Run only when you are logged in"
Write-Host "  - Run using your current Windows account"
Write-Host "  - Prevent multiple copies running simultaneously"
Write-Host "  - Retry automatically if the task fails"
Write-Host ""


# ============================================================
# ASK FOR SCRIPT PATH
# ============================================================

while ($true) {

    $scriptPath = Read-Host `
        "Enter the full path to the PowerShell script to schedule"

    $scriptPath = $scriptPath.Trim('"').Trim()


    if ([string]::IsNullOrWhiteSpace($scriptPath)) {

        Write-Host ""
        Write-Host "Please enter a script path."
        Write-Host ""

        continue
    }


    if (-not (Test-Path $scriptPath -PathType Leaf)) {

        Write-Host ""
        Write-Host "ERROR: File not found:"
        Write-Host $scriptPath
        Write-Host ""

        continue
    }


    if (
        [System.IO.Path]::GetExtension($scriptPath) -ne ".ps1"
    ) {

        Write-Host ""
        Write-Host "ERROR: The selected file does not appear to be a .ps1 file."
        Write-Host ""

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
# ASK FOR INTERVAL
# ============================================================

Write-Host ""

while ($true) {

    $intervalInput = Read-Host `
        "Run every how many minutes? [5]"

    if ([string]::IsNullOrWhiteSpace($intervalInput)) {

        $intervalMinutes = 5

        break
    }


    if (
        [int]::TryParse(
            $intervalInput,
            [ref]$intervalMinutes
        )
    ) {

        if ($intervalMinutes -ge 1) {

            break
        }
    }


    Write-Host ""
    Write-Host "Please enter a whole number of at least 1."
    Write-Host ""
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

Write-Host "Interval:"
Write-Host "  Every $intervalMinutes minute(s)"

Write-Host ""

Write-Host "User:"
Write-Host "  $env:USERDOMAIN\$env:USERNAME"

Write-Host ""

Write-Host "Run mode:"
Write-Host "  Only when user is logged in"

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
# REMOVE EXISTING TASK IF PRESENT
# ============================================================

$existingTask = Get-ScheduledTask `
    -TaskName $taskName `
    -ErrorAction SilentlyContinue


if ($existingTask) {

    Write-Host ""
    Write-Host "A task with this name already exists."

    $replace = Read-Host `
        "Replace it? [y/N]"

    if (
        $replace -match "^(Y|y|Yes|yes)$"
    ) {

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
# POWERSHELL PATH
# ============================================================

$powerShellPath = Join-Path `
    $env:SystemRoot `
    "System32\WindowsPowerShell\v1.0\powershell.exe"


if (-not (Test-Path $powerShellPath)) {

    throw "Windows PowerShell executable was not found."
}


# ============================================================
# ACTION
# ============================================================

$action = New-ScheduledTaskAction `
    -Execute $powerShellPath `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""


# ============================================================
# TRIGGER
#
# Start 1 minute from now.
#
# Then repeat at the selected interval indefinitely.
# ============================================================

$startTime = (Get-Date).AddMinutes(1)

$trigger = New-ScheduledTaskTrigger `
    -Once `
    -At $startTime `
    -RepetitionInterval (
        New-TimeSpan -Minutes $intervalMinutes
    ) `
    -RepetitionDuration (
        New-TimeSpan -Days 3650
    )


# ============================================================
# PRINCIPAL
#
# InteractiveToken means:
#
#   Run only when user is logged on
#
# This is important for Outlook COM.
# ============================================================

$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive `
    -RunLevel Limited


# ============================================================
# SETTINGS
# ============================================================

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -RestartCount 3 `
    -RestartInterval (
        New-TimeSpan -Minutes 5
    )


# ============================================================
# CREATE TASK
# ============================================================

Write-Host ""
Write-Host "Creating scheduled task..."

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Automatically runs the Worklog Outlook ICS importer."


# ============================================================
# SUCCESS
# ============================================================

Write-Host ""
Write-Host "=============================================="
Write-Host " SUCCESS"
Write-Host "=============================================="
Write-Host ""

Write-Host "Scheduled task created:"
Write-Host "  $taskName"

Write-Host ""

Write-Host "Script:"
Write-Host "  $scriptPath"

Write-Host ""

Write-Host "Schedule:"
Write-Host "  Every $intervalMinutes minute(s)"

Write-Host ""

Write-Host "The first run will occur around:"
Write-Host "  $startTime"

Write-Host ""

Write-Host "It will run only while you are logged into Windows."

Write-Host ""

# ============================================================
# ASK WHETHER TO RUN NOW
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

    Write-Host "The PowerShell script should now be running."
}

Write-Host ""
Write-Host "Setup complete."
Write-Host ""