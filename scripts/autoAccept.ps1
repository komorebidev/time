# ============================================================
# Worklog ICS -> Outlook Calendar
#
# Reads the latest "Work Log ICS: time" email from Outlook,
# extracts time.ics, and creates/updates all calendar events.
#
# Designed for classic Outlook desktop + Outlook COM.
# ============================================================

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------

$SubjectPrefix = "Work Log ICS:"

# Temporary location for the downloaded ICS
$TempICS = Join-Path $env:TEMP "worklog-time.ics"


# ------------------------------------------------------------
# ICS TEXT DECODING
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# ICS DATE PARSER
#
# Your Python currently generates:
#
# DTSTART;VALUE=DATE:20260805
# DTEND;VALUE=DATE:20260806
#
# This also supports basic date/time formats.
# ------------------------------------------------------------

function Parse-IcsDate {

    param(
        [string]$Value
    )

    if ($Value -match '^(\d{4})(\d{2})(\d{2})$') {

        return [datetime]::ParseExact(
            $Value,
            "yyyyMMdd",
            $null
        )
    }

    if ($Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$') {

        return [datetime]::ParseExact(
            $Value,
            "yyyyMMddTHHmmssZ",
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal
        ).ToLocalTime()
    }

    if ($Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})$') {

        return [datetime]::ParseExact(
            $Value,
            "yyyyMMddTHHmmss",
            $null
        )
    }

    throw "Unsupported ICS date format: $Value"
}


# ------------------------------------------------------------
# PARSE ICS FILE
# ------------------------------------------------------------

function Read-IcsEvents {

    param(
        [string]$Path
    )

    $raw = [System.IO.File]::ReadAllText(
        $Path,
        [System.Text.Encoding]::UTF8
    )

    # Normalize line endings
    $raw = $raw -replace "`r`n", "`n"
    $raw = $raw -replace "`r", "`n"

    # --------------------------------------------------------
    # Unfold iCalendar continuation lines
    # --------------------------------------------------------

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

        $propertyPart = $parts[0]
        $value = $parts[1]

        # Remove parameters such as:
        # DTSTART;VALUE=DATE
        $propertyName = ($propertyPart -split ";")[0].ToUpper()

        $current.Properties[$propertyName] = $value
    }

    return $events
}


# ------------------------------------------------------------
# FIND EXISTING WORKLOG EVENT
#
# We use a custom Outlook UserProperty called:
#
#     WorklogUID
#
# This lets us reliably update an existing event instead of
# creating duplicates.
# ------------------------------------------------------------

function Find-ExistingWorklogEvent {

    param(
        $Calendar,
        [string]$UID
    )

    $items = $null

    try {

        $items = $Calendar.Items

        foreach ($item in $items) {

            # Only appointment items
            if ($item.Class -ne 26) {
                continue
            }

            $property = $null

            try {

                $property = $item.UserProperties.Find(
                    "WorklogUID"
                )

                if (
                    $null -ne $property -and
                    $property.Value -eq $UID
                ) {

                    return $item
                }
            }
            catch {
                # Ignore malformed/non-standard items
            }

            if ($property) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $property
                ) | Out-Null
            }
        }
    }
    finally {

        if ($items) {

            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $items
            ) | Out-Null
        }
    }

    return $null
}


# ============================================================
# START OUTLOOK
# ============================================================

$outlook = $null
$namespace = $null
$calendar = $null

try {

    Write-Host ""
    Write-Host "========================================"
    Write-Host "Worklog Outlook Importer"
    Write-Host "========================================"
    Write-Host ""

    Write-Host "Connecting to Outlook..."

    $outlook = New-Object -ComObject Outlook.Application

    $namespace = $outlook.GetNamespace("MAPI")

    Write-Host "✓ Outlook connected"

    # --------------------------------------------------------
    # Default Inbox
    # 6 = olFolderInbox
    # --------------------------------------------------------

    $inbox = $namespace.GetDefaultFolder(6)

    Write-Host "✓ Inbox accessed"

    # --------------------------------------------------------
    # Default Calendar
    #
    # 9 = olFolderCalendar
    # --------------------------------------------------------

    $calendar = $namespace.GetDefaultFolder(9)

    Write-Host "✓ Calendar accessed"
    Write-Host ""


    # ========================================================
    # FIND LATEST WORKLOG EMAIL
    # ========================================================

    Write-Host "Searching for latest Work Log email..."

    $mailItems = $inbox.Items

    # Sort newest first
    $mailItems.Sort(
        "[ReceivedTime]",
        $true
    )

    $targetMail = $null

    foreach ($mail in $mailItems) {

        if ($mail.Class -ne 43) {
            continue
        }

        $subject = [string]$mail.Subject

        if (
            $subject.StartsWith(
                $SubjectPrefix,
                [StringComparison]::OrdinalIgnoreCase
            )
        ) {

            if ($mail.Attachments.Count -gt 0) {

                $targetMail = $mail

                break
            }
        }
    }


    if ($null -eq $targetMail) {

        Write-Host "No Work Log ICS email found."
        exit 0
    }


    Write-Host "✓ Found:"
    Write-Host "  $($targetMail.Subject)"
    Write-Host "  Received: $($targetMail.ReceivedTime)"
    Write-Host ""


    # ========================================================
    # FIND time.ics
    # ========================================================

    $icsAttachment = $null

    for ($i = 1; $i -le $targetMail.Attachments.Count; $i++) {

        $attachment = $targetMail.Attachments.Item($i)

        if (
            $attachment.FileName -ieq "time.ics"
        ) {

            $icsAttachment = $attachment

            break
        }
    }


    if ($null -eq $icsAttachment) {

        Write-Host "time.ics attachment not found."

        exit 1
    }


    Write-Host "✓ Found attachment: time.ics"


    # ========================================================
    # SAVE ICS TEMPORARILY
    # ========================================================

    if (Test-Path $TempICS) {

        Remove-Item $TempICS -Force
    }


    $icsAttachment.SaveAsFile(
        $TempICS
    )

    Write-Host "✓ Saved ICS temporarily"
    Write-Host ""


    # ========================================================
    # PARSE ICS
    # ========================================================

    $events = Read-IcsEvents $TempICS

    Write-Host "Events found: $($events.Count)"
    Write-Host ""


    # ========================================================
    # IMPORT EACH EVENT
    # ========================================================

    $createdCount = 0
    $updatedCount = 0
    $skippedCount = 0


    foreach ($icsEvent in $events) {

        $properties = $icsEvent.Properties

        $uid = $properties["UID"]

        if ([string]::IsNullOrWhiteSpace($uid)) {

            Write-Host "Skipping event without UID."

            $skippedCount++

            continue
        }


        $summary = Unescape-IcsText(
            $properties["SUMMARY"]
        )

        $description = Unescape-IcsText(
            $properties["DESCRIPTION"]
        )

        $location = Unescape-IcsText(
            $properties["LOCATION"]
        )


        try {

            $start = Parse-IcsDate(
                $properties["DTSTART"]
            )

            $end = Parse-IcsDate(
                $properties["DTEND"]
            )

        }
        catch {

            Write-Host "Could not parse dates for $uid"
            Write-Host $_.Exception.Message

            $skippedCount++

            continue
        }


        Write-Host "Processing:"
        Write-Host "  UID: $uid"
        Write-Host "  Subject: $summary"
        Write-Host "  Date: $($start.ToString('yyyy-MM-dd'))"


        # ----------------------------------------------------
        # Find existing event
        # ----------------------------------------------------

        $existing = Find-ExistingWorklogEvent `
            -Calendar $calendar `
            -UID $uid


        # ----------------------------------------------------
        # Create new event
        # ----------------------------------------------------

        if ($null -eq $existing) {

            $appointment = $calendar.Items.Add(1)

            # Subject
            $appointment.Subject = $summary

            # All-day
            $appointment.AllDayEvent = $true

            # Dates
            $appointment.Start = $start
            $appointment.End = $end

            # Description
            $appointment.Body = $description

            # Location
            $appointment.Location = $location


            # ------------------------------------------------
            # Out of Office
            #
            # 3 = olOutOfOffice
            # ------------------------------------------------

            if (
                $summary -eq "Out of Office"
            ) {

                $appointment.BusyStatus = 3
            }
            else {

                $appointment.BusyStatus = 2
            }


            # ------------------------------------------------
            # Store our UID inside Outlook
            # ------------------------------------------------

            $uidProperty =
                $appointment.UserProperties.Add(
                    "WorklogUID",
                    1,
                    $false
                )

            $uidProperty.Value = $uid

            $appointment.Save()


            Write-Host "  ✓ Created"

            $createdCount++


            # Release objects
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $uidProperty
            ) | Out-Null

            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $appointment
            ) | Out-Null

        }

        # ----------------------------------------------------
        # Update existing event
        # ----------------------------------------------------

        else {

            $existing.Subject = $summary

            $existing.AllDayEvent = $true

            $existing.Start = $start
            $existing.End = $end

            $existing.Body = $description

            $existing.Location = $location


            if (
                $summary -eq "Out of Office"
            ) {

                $existing.BusyStatus = 3
            }
            else {

                $existing.BusyStatus = 2
            }


            $existing.Save()


            Write-Host "  ✓ Updated"

            $updatedCount++


            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $existing
            ) | Out-Null
        }


        Write-Host ""
    }


    # ========================================================
    # SUMMARY
    # ========================================================

    Write-Host "========================================"
    Write-Host "IMPORT COMPLETE"
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Created: $createdCount"
    Write-Host "Updated: $updatedCount"
    Write-Host "Skipped: $skippedCount"
    Write-Host ""


}
catch {

    Write-Host ""
    Write-Host "========================================"
    Write-Host "ERROR"
    Write-Host "========================================"
    Write-Host ""
    Write-Host $_.Exception.Message
    Write-Host ""
    Write-Host $_.ScriptStackTrace

    exit 1
}
finally {

    if (Test-Path $TempICS) {

        Remove-Item $TempICS -Force -ErrorAction SilentlyContinue
    }


    if ($calendar) {

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $calendar
        ) | Out-Null
    }

    if ($namespace) {

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $namespace
        ) | Out-Null
    }

    if ($outlook) {

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $outlook
        ) | Out-Null
    }


    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}