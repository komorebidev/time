# ============================================================
# Worklog ICS -> Outlook Calendar
#
# Uses classic Outlook COM/MAPI.
#
# Handles:
#   - Work Log ICS emails
#   - time.ics attachments
#   - Multiple VEVENT entries
#   - Duplicate prevention
#   - Calendar updates
#   - Outlook startup timing problems
#   - Outlook restart minimized (forced via API)
#   - Safe Outlook shutdown upon completion
#
# ============================================================

$ErrorActionPreference = "Stop"

# ============================================================
# WIN32 API FOR FORCED MINIMIZATION
# ============================================================

$definition = @'
using System;
using System.Runtime.InteropServices;

public class WindowHelper {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
'@

Add-Type -TypeDefinition $definition -ErrorAction SilentlyContinue

# ============================================================
# CONFIGURATION
# ============================================================

$SubjectSearch = "Work Log ICS:"
$AttachmentName = "time.ics"
$ProcessedCategory = "WorkLog Imported"
$TempICS = Join-Path $env:TEMP "worklog-time.ics"

# ============================================================
# OUTLOOK OWNERSHIP TRACKING
# ============================================================

$script:StartedOutlookByScript = $false

# ============================================================
# CLOSE EXISTING OUTLOOK
# ============================================================

function Close-Outlook {
    Write-Host ""
    Write-Host "Checking existing Outlook process..."

    $process = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue

    if ($null -eq $process) {
        Write-Host "No Outlook process found."
        return
    }

    Write-Host "Existing Outlook detected."
    Write-Host "Closing Outlook..."

    try {
        $process | Stop-Process -Force
        Start-Sleep -Seconds 5
        Write-Host "Outlook closed."
    }
    catch {
        Write-Host "Failed to close Outlook."
        throw
    }
}

# ============================================================
# START OUTLOOK MINIMIZED (WITH FORCED API MINIMIZE)
# ============================================================

function Start-Outlook-Minimized {
    Write-Host ""
    Write-Host "Starting Outlook minimized..."

    Start-Process "outlook.exe"
    $script:StartedOutlookByScript = $true

    # Give Outlook a few moments to spawn its main window handle, then force minimize it
    Start-Sleep -Seconds 2

    $swMinimized = 2 # SW_SHOWMINIMIZED
    $maxTries = 10
    
    for ($i = 0; $i -lt $maxTries; $i++) {
        $outlookProc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
        if ($outlookProc -and $outlookProc.MainWindowHandle -ne [IntPtr]::Zero) {
            [WindowHelper]::ShowWindow($outlookProc.MainWindowHandle, $swMinimized) | Out-Null
            break
        }
        Start-Sleep -Seconds 1
    }

    Start-Sleep -Seconds 3
}

# ============================================================
# CONNECT TO OUTLOOK
# ============================================================

function Connect-Outlook {
    $maxAttempts = 30
    $delaySeconds = 5

    Close-Outlook
    Start-Outlook-Minimized

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            Write-Host ""
            Write-Host "Connecting Outlook COM..."
            Write-Host "Attempt $attempt of $maxAttempts"

            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            Write-Host "✓ Outlook COM connected"

            $namespace = $outlook.GetNamespace("MAPI")
            $inbox = $namespace.GetDefaultFolder(6)

            if ($null -eq $inbox) {
                throw "Inbox unavailable."
            }

            Write-Host "✓ MAPI ready"

            [Runtime.InteropServices.Marshal]::ReleaseComObject($inbox) | Out-Null
            [Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null

            return $outlook
        }
        catch {
            Write-Host ""
            Write-Host "Outlook COM not ready."
            Write-Host $_.Exception.Message

            if ($_.Exception.HResult -eq -2147418111) {
                Write-Host "RPC_E_CALL_REJECTED detected."
            }

            if ($_.Exception.HResult -eq -2147221021) {
                Write-Host "MK_E_UNAVAILABLE detected."
            }

            if ($attempt -lt $maxAttempts) {
                Start-Sleep -Seconds $delaySeconds
            }
            else {
                throw
            }
        }
    }
}

# ============================================================
# ICS TEXT UNESCAPING
# ============================================================

function Unescape-IcsText {
    param(
        [string]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    $Text = $Text -replace '\\n', "`r`n"
    $Text = $Text -replace '\\N', "`r`n"
    $Text = $Text -replace '\\,', ","
    $Text = $Text -replace '\\;', ";"
    $Text = $Text -replace '\\\\', "\"
    return $Text
}

# ============================================================
# ICS DATE PARSER
# ============================================================

function Parse-IcsDate {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "ICS date value is empty."
    }

    if ($Value -match '^(\d{4})(\d{2})(\d{2})$') {
        return [datetime]::ParseExact(
            $Value,
            "yyyyMMdd",
            [Globalization.CultureInfo]::InvariantCulture
        )
    }

    if ($Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$') {
        return (
            [datetime]::ParseExact(
                $Value,
                "yyyyMMddTHHmmssZ",
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::AssumeUniversal
            )
        ).ToLocalTime()
    }

    if ($Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})$') {
        return [datetime]::ParseExact(
            $Value,
            "yyyyMMddTHHmmss",
            [Globalization.CultureInfo]::InvariantCulture
        )
    }

    throw "Unsupported ICS date format: $Value"
}

# ============================================================
# READ ICS EVENTS
# ============================================================

function Read-IcsEvents {
    param(
        [string]$Path
    )

    $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)

    # Normalize line endings
    $raw = $raw -replace "`r`n", "`n"
    $raw = $raw -replace "`r", "`n"

    # iCalendar unfolding
    $raw = $raw -replace "`n[ `t]", ""

    $lines = $raw -split "`n"
    $events = @()
    $current = $null

    foreach ($line in $lines) {
        $line = $line.TrimEnd()

        if ($line -eq "BEGIN:VEVENT") {
            $current = @{
                Properties = @{}
            }
            continue
        }

        if ($line -eq "END:VEVENT") {
            if ($null -ne $current) {
                $events += $current
            }
            $current = $null
            continue
        }

        if ($null -eq $current) {
            continue
        }

        $parts = $line -split ":", 2

        if ($parts.Count -ne 2) {
            continue
        }

        $propertyName = ($parts[0] -split ";")[0].ToUpper()
        $value = $parts[1]

        $current.Properties[$propertyName] = $value
    }

    return @($events)
}

# ============================================================
# FIND EXISTING WORKLOG EVENT
# ============================================================

function Find-ExistingWorklogEvent {
    param(
        $Calendar,
        [string]$UID
    )

    $items = $null

    try {
        $items = $Calendar.Items
        $count = $items.Count

        for ($index = 1; $index -le $count; $index++) {
            $item = $null

            try {
                $item = $items.Item($index)

                if ($item.Class -ne 26) {
                    continue
                }

                $property = $null

                try {
                    $property = $item.UserProperties.Find("WorklogUID")

                    if ($null -ne $property) {
                        if ($property.Value -eq $UID) {
                            return $item
                        }
                    }
                }
                finally {
                    if ($property) {
                        [Runtime.InteropServices.Marshal]::ReleaseComObject($property) | Out-Null
                    }
                }
            }
            catch {
                # Ignore invalid calendar items
            }
        }
    }
    finally {
        if ($items) {
            [Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null
        }
    }

    return $null
}

# ============================================================
# FIND WORKLOG EMAILS
# ============================================================

function Find-WorklogEmails {
    param(
        $Inbox
    )

    $matches = @()
    $items = $Inbox.Items

    try {
        $items.Sort("[ReceivedTime]", $true)
        $count = $items.Count

        for ($index = 1; $index -le $count; $index++) {
            $mail = $null

            try {
                $mail = $items.Item($index)

                if ($mail.Class -ne 43) {
                    continue
                }

                $subject = [string]$mail.Subject

                if ($subject -notlike "*$SubjectSearch*") {
                    continue
                }

                $hasAttachment = $false
                $attachmentCount = $mail.Attachments.Count

                for ($a = 1; $a -le $attachmentCount; $a++) {
                    $attachment = $null

                    try {
                        $attachment = $mail.Attachments.Item($a)

                        if ($attachment.FileName -ieq $AttachmentName) {
                            $hasAttachment = $true
                            break
                        }
                    }
                    finally {
                        if ($attachment) {
                            [Runtime.InteropServices.Marshal]::ReleaseComObject($attachment) | Out-Null
                        }
                    }
                }

                if ($hasAttachment) {
                    $matches += $mail
                }
                else {
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null
                }
            }
            catch {
                if ($mail) {
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null
                }
            }
        }
    }
    finally {
        if ($items) {
            [Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null
        }
    }

    return @($matches)
}

# ============================================================
# MAIN VARIABLES
# ============================================================

$outlook = $null
$namespace = $null
$inbox = $null
$calendar = $null
$targetMail = $null
$icsAttachment = $null
$createdCount = 0
$updatedCount = 0
$skippedCount = 0

# ============================================================
# MAIN
# ============================================================

try {
    Write-Host ""
    Write-Host "========================================"
    Write-Host "WORKLOG OUTLOOK IMPORTER"
    Write-Host "========================================"
    Write-Host ""

    # CONNECT OUTLOOK
    $outlook = Connect-Outlook

    # LOAD PROFILE
    Write-Host ""
    Write-Host "Loading Outlook profile..."

    $namespace = $outlook.GetNamespace("MAPI")
    Write-Host "✓ MAPI loaded"

    # GET FOLDERS
    $inbox = $namespace.GetDefaultFolder(6)
    $calendar = $namespace.GetDefaultFolder(9)

    Write-Host ""
    Write-Host "Inbox:"
    Write-Host $inbox.FolderPath

    Write-Host ""
    Write-Host "Calendar:"
    Write-Host $calendar.FolderPath

    # SEARCH EMAILS
    Write-Host ""
    Write-Host "Searching Work Log ICS emails..."

    $matchingEmails = @(Find-WorklogEmails -Inbox $inbox)

    if ($matchingEmails.Count -eq 0) {
        Write-Host ""
        Write-Host "No Work Log ICS emails found."
        exit 0
    }

    # Newest first
    $targetMail = $matchingEmails[0]

    Write-Host ""
    Write-Host "Selected email:"
    Write-Host $targetMail.Subject
    Write-Host $targetMail.ReceivedTime

    # DELETE OLDER EMAILS
    if ($matchingEmails.Count -gt 1) {
        Write-Host ""
        Write-Host "Deleting older matching emails..."

        for ($i = 1; $i -lt $matchingEmails.Count; $i++) {
            try {
                $matchingEmails[$i].Delete()
            }
            catch {
                Write-Host "Unable to delete older email."
            }
        }

        Write-Host "✓ Old emails deleted"
    }

    # FIND time.ics
    $attachmentCount = $targetMail.Attachments.Count

    for ($i = 1; $i -le $attachmentCount; $i++) {
        $attachment = $targetMail.Attachments.Item($i)

        if ($attachment.FileName -ieq $AttachmentName) {
            $icsAttachment = $attachment
            break
        }

        [Runtime.InteropServices.Marshal]::ReleaseComObject($attachment) | Out-Null
    }

    if ($null -eq $icsAttachment) {
        throw "time.ics attachment not found."
    }

    # SAVE ICS FILE
    if (Test-Path $TempICS) {
        Remove-Item $TempICS -Force -ErrorAction SilentlyContinue
    }

    Write-Host ""
    Write-Host "Saving ICS..."

    $icsAttachment.SaveAsFile($TempICS)

    Write-Host "✓ ICS saved"

    # PARSE EVENTS
    $events = Read-IcsEvents -Path $TempICS

    Write-Host ""
    Write-Host "Events found:"
    Write-Host $events.Count

    # ============================================================
    # PROCESS EVENTS
    # ============================================================

    foreach ($icsEvent in $events) {
        $properties = $icsEvent.Properties

        # UID
        $uid = $properties["UID"]

        if ([string]::IsNullOrWhiteSpace($uid)) {
            Write-Host ""
            Write-Host "Skipping event without UID."
            $skippedCount++
            continue
        }

        # TEXT FIELDS
        $summary = Unescape-IcsText($properties["SUMMARY"])
        $description = Unescape-IcsText($properties["DESCRIPTION"])
        $location = Unescape-IcsText($properties["LOCATION"])

        # DATES
        try {
            $start = Parse-IcsDate($properties["DTSTART"])
            $end = Parse-IcsDate($properties["DTEND"])
        }
        catch {
            Write-Host ""
            Write-Host "Invalid event date."
            Write-Host $uid
            $skippedCount++
            continue
        }

        Write-Host ""
        Write-Host "--------------------------------"
        Write-Host "Processing:"
        Write-Host $summary
        Write-Host "UID:"
        Write-Host $uid

        # LOOK FOR EXISTING EVENT
        $existing = Find-ExistingWorklogEvent -Calendar $calendar -UID $uid

        # CREATE EVENT
        if ($null -eq $existing) {
            Write-Host "Action: CREATE"

            $appointment = $null
            $uidProperty = $null

            try {
                $appointment = $calendar.Items.Add(1)

                $appointment.Subject = $summary
                $appointment.Start = $start
                $appointment.End = $end
                $appointment.AllDayEvent = $true
                $appointment.Body = $description
                $appointment.Location = $location

                if ($summary -eq "Out of Office") {
                    $appointment.BusyStatus = 3
                }
                else {
                    $appointment.BusyStatus = 2
                }

                $uidProperty = $appointment.UserProperties.Add("WorklogUID", 1, $false)
                $uidProperty.Value = $uid

                $appointment.Save()

                Write-Host "✓ Created"
                $createdCount++
            }
            finally {
                if ($uidProperty) {
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($uidProperty) | Out-Null
                }

                if ($appointment) {
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($appointment) | Out-Null
                }
            }
        }
        else {
            Write-Host "Action: ALREADY EXISTS (Skipping creation)"
            $skippedCount++
        }
    }
}
finally {
    # Cleanup main COM objects
    if ($icsAttachment) { [Runtime.InteropServices.Marshal]::ReleaseComObject($icsAttachment) | Out-Null }
    if ($targetMail) { [Runtime.InteropServices.Marshal]::ReleaseComObject($targetMail) | Out-Null }
    if ($calendar) { [Runtime.InteropServices.Marshal]::ReleaseComObject($calendar) | Out-Null }
    if ($inbox) { [Runtime.InteropServices.Marshal]::ReleaseComObject($inbox) | Out-Null }
    if ($namespace) { [Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null }
    if ($outlook) { [Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Clean up temp file
    if (Test-Path $TempICS) {
        Remove-Item $TempICS -Force -ErrorAction SilentlyContinue
    }

    # Close Outlook since the import is complete (as requested)
    if ($script:StartedOutlookByScript) {
        Write-Host ""
        Write-Host "Closing Outlook since import is finished..."
        Close-Outlook
    }

    Write-Host ""
    Write-Host "========================================"
    Write-Host "FINISHED"
    Write-Host "Created: $createdCount | Skipped/Existing: $skippedCount"
    Write-Host "========================================"
}