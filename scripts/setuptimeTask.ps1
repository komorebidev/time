# ============================================================
# Interactive Scheduled Task Setup
#
# Creates a Windows Scheduled Task that runs a PowerShell
# script whenever the workstation is unlocked.
#
# Features:
#   - Interactive script-path prompt
#   - Interactive task-name prompt
#   - Runs PowerShell hidden
#   - Runs only when the current user is logged in
#   - Prevents overlapping executions
#   - Starts when the workstation is unlocked
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
Write-Host "  When the workstation is unlocked"

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
#
# We intentionally use Windows PowerShell 5.1 because the
# target script uses Outlook COM.
# ============================================================

$powerShellPath = Join-Path `
    $env:SystemRoot `
    "System32\WindowsPowerShell\v1.0\powershell.exe"


if (-not (Test-Path -LiteralPath $powerShellPath)) {

    throw "Windows PowerShell executable was not found."
}


# ============================================================
# CREATE ACTION
#
# -NoProfile
#       Avoids loading the user's PowerShell profile.
#
# -WindowStyle Hidden
#       Prevents a PowerShell console window appearing.
#
# -ExecutionPolicy Bypass
#       Allows the selected script to execute.
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
# CREATE PRINCIPAL
#
# InteractiveToken means:
#
#   Run only when the current user is logged in.
#
# This is important for Outlook COM automation.
# ============================================================

$currentUser = `
    "$env:USERDOMAIN\$env:USERNAME"


$principal = New-ScheduledTaskPrincipal `
    -UserId $currentUser `
    -LogonType Interactive `
    -RunLevel Limited


# ============================================================
# CREATE SETTINGS
#
# AllowStartIfOnBatteries
#   Allows the task to run on battery.
#
# DontStopIfGoingOnBatteries
#   Doesn't stop an active run when switching to battery.
#
# StartWhenAvailable
#   Allows Windows to start the task when available.
#
# MultipleInstances IgnoreNew
#   Prevents overlapping executions.
#
# RestartCount / RestartInterval
#   Retries failed task executions.
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
# CREATE TASK XML
#
# Event trigger:
#
#   Security log
#   Event ID 4801
#
# Event 4801 means:
#   "The workstation was unlocked."
#
# The XML is used because the ScheduledTasks PowerShell
# cmdlets do not expose a simple workstation-unlock trigger.
# ============================================================

$taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4"
      xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">

  <RegistrationInfo>
    <Description>
      Automatically runs the Worklog Outlook ICS importer
      whenever the workstation is unlocked.
    </Description>
  </RegistrationInfo>

  <Triggers>

    <EventTrigger>

      <Enabled>true</Enabled>

      <Subscription>
        <![CDATA[
        <QueryList>
          <Query Id="0" Path="Security">
            <Select Path="Security">
              *[System[
                Provider[@Name='Microsoft-Windows-Security-Auditing']
                and (EventID=4801)
              ]]
            </Select>
          </Query>
        </QueryList>
        ]]>
      </Subscription>

    </EventTrigger>

  </Triggers>

  <Principals>

    <Principal id="Author">

      <UserId>$currentUser</UserId>

      <LogonType>InteractiveToken</LogonType>

      <RunLevel>LeastPrivilege</RunLevel>

    </Principal>

  </Principals>

  <Settings>

    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>

    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>

    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>

    <StartWhenAvailable>true</StartWhenAvailable>

    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>

    <RestartOnFailure>

      <Interval>PT5M</Interval>

      <Count>3</Count>

    </RestartOnFailure>

  </Settings>

  <Actions Context="Author">

    <Exec>

      <Command>$powerShellPath</Command>

      <Arguments>$arguments</Arguments>

    </Exec>

  </Actions>

</Task>
"@


# ============================================================
# REGISTER TASK
# ============================================================

Write-Host ""
Write-Host "Creating scheduled task..."

$taskXml |
    Register-ScheduledTask `
        -TaskName $taskName `
        -Force


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
Write-Host "  When the workstation is unlocked"

Write-Host ""

Write-Host "PowerShell window:"
Write-Host "  Hidden"

Write-Host ""

Write-Host "Run mode:"
Write-Host "  Only while you are logged in"

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

    Write-Host ""
    Write-Host "You can check Task Scheduler to verify"
    Write-Host "the Last Run Result if necessary."
}


# ============================================================
# DONE
# ============================================================

Write-Host ""
Write-Host "Setup complete."