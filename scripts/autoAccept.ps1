# ============================================================
# Worklog ICS -> Outlook Calendar
#
# Reads the latest email containing:
#
#     Work Log ICS:
#
# and a time.ics attachment.
#
# Supports forwarded emails such as:
#
#     FW: Work Log ICS: time
#
# Uses classic Outlook COM/MAPI.
#
# Handles:
#   - Multiple VEVENTs in one ICS
#   - All-day events
#   - Descriptions
#   - Locations
#   - Out Of Office events
#   - Duplicate prevention
#   - Updating existing events
#   - Outlook startup/RPC timing problems
# ============================================================


$ErrorActionPreference = "Stop"


# ============================================================
# CONFIGURATION
# ============================================================

$SubjectSearch = "Work Log ICS:"

$AttachmentName = "time.ics"

$TempICS = Join-Path `
    $env:TEMP `
    "worklog-time.ics"


# ============================================================
# CONNECT TO OUTLOOK WITH RETRIES
#
# Outlook can temporarily reject COM requests while starting
# or loading the profile.
#
# RPC_E_CALL_REJECTED
# 0x80010001
#
# We retry instead of failing immediately.
# ============================================================

function Connect-Outlook {

    $maxAttempts = 12

    $delaySeconds = 5

    for (
        $attempt = 1;
        $attempt -le $maxAttempts;
        $attempt++
    ) {

        try {

            Write-Host ""
            Write-Host "Connecting to Outlook..."
            Write-Host "Attempt $attempt of $maxAttempts"

            $outlook = New-Object `
                -ComObject Outlook.Application

            Write-Host "✓ Outlook COM connected"

            return $outlook
        }
        catch {

            $message = $_.Exception.Message

            Write-Host ""
            Write-Host "Outlook is not ready yet."
            Write-Host $message

            if ($attempt -lt $maxAttempts) {

                Write-Host ""
                Write-Host "Waiting $delaySeconds seconds..."
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
#
# Supports:
#
#     20260805
#     20260805T120000
#     20260805T120000Z
# ============================================================

function Parse-IcsDate {

    param(
        [string]$Value
    )


    if ([string]::IsNullOrWhiteSpace($Value)) {

        throw "ICS date value is empty."
    }


    # --------------------------------------------------------
    # Date only
    #
    # Example:
    #
    # 20260805
    # --------------------------------------------------------

    if (
        $Value -match
        '^(\d{4})(\d{2})(\d{2})$'
    ) {

        return [datetime]::ParseExact(
            $Value,
            "yyyyMMdd",
            [Globalization.CultureInfo]::InvariantCulture
        )
    }


    # --------------------------------------------------------
    # UTC date/time
    #
    # Example:
    #
    # 20260805T120000Z
    # --------------------------------------------------------

    if (
        $Value -match
        '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$'
    ) {

        return [datetime]::ParseExact(
            $Value,
            "yyyyMMddTHHmmssZ",
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal
        ).ToLocalTime()
    }


    # --------------------------------------------------------
    # Local date/time
    #
    # Example:
    #
    # 20260805T120000
    # --------------------------------------------------------

    if (
        $Value -match
        '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})$'
    ) {

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


    $raw = [System.IO.File]::ReadAllText(
        $Path,
        [System.Text.Encoding]::UTF8
    )


    # --------------------------------------------------------
    # Normalize line endings
    # --------------------------------------------------------

    $raw = $raw -replace "`r`n", "`n"

    $raw = $raw -replace "`r", "`n"


    # --------------------------------------------------------
    # iCalendar line unfolding
    #
    # A line beginning with space/tab continues the previous
    # line.
    # --------------------------------------------------------

    $raw = $raw -replace "`n[ `t]", ""


    $lines = $raw -split "`n"


    $events = @()

    $current = $null


    foreach ($line in $lines) {

        $line = $line.TrimEnd()


        # ----------------------------------------------------
        # BEGIN:VEVENT
        # ----------------------------------------------------

        if ($line -eq "BEGIN:VEVENT") {

            $current = @{
                Properties = @{}
            }

            continue
        }


        # ----------------------------------------------------
        # END:VEVENT
        # ----------------------------------------------------

        if ($line -eq "END:VEVENT") {

            if ($null -ne $current) {

                $events += $current
            }

            $current = $null

            continue
        }


        # ----------------------------------------------------
        # Ignore everything outside VEVENT
        # ----------------------------------------------------

        if ($null -eq $current) {

            continue
        }


        # ----------------------------------------------------
        # Split property from value
        #
        # Example:
        #
        # DTSTART;VALUE=DATE:20260805
        #
        # becomes:
        #
        # property = DTSTART;VALUE=DATE
        # value    = 20260805
        # ----------------------------------------------------

        $parts = $line -split ":", 2


        if ($parts.Count -ne 2) {

            continue
        }


        $propertyPart = $parts[0]

        $value = $parts[1]


        # ----------------------------------------------------
        # Remove parameters
        #
        # DTSTART;VALUE=DATE
        #
        # becomes:
        #
        # DTSTART
        # ----------------------------------------------------

        $propertyName =
            ($propertyPart -split ";")[0].ToUpper()


        $current.Properties[$propertyName] = $value
    }


    return $events
}


# ============================================================
# FIND EXISTING WORKLOG EVENT
#
# We store the ICS UID in an Outlook custom property:
#
#     WorklogUID
#
# This allows the script to safely run repeatedly without
# creating duplicates.
# ============================================================

function Find-ExistingWorklogEvent {

    param(
        $Calendar,
        [string]$UID
    )


    $items = $null


    try {

        $items = $Calendar.Items


        foreach ($item in $items) {

            try {

                # 26 = Outlook AppointmentItem
                if ($item.Class -ne 26) {

                    continue
                }


                $property = $null


                try {

                    $property =
                        $item.UserProperties.Find(
                            "WorklogUID"
                        )


                    if (
                        $null -ne $property -and
                        $property.Value -eq $UID
                    ) {

                        return $item
                    }
                }
                finally {

                    if ($property) {

                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                            $property
                        ) | Out-Null
                    }
                }
            }
            catch {

                # Ignore individual malformed items
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
# MAIN
# ============================================================

$outlook = $null

$namespace = $null

$inbox = $null

$calendar = $null

$mailItems = $null

$targetMail = $null

$icsAttachment = $null


try {

    Write-Host ""
    Write-Host "========================================"
    Write-Host "WORKLOG OUTLOOK IMPORTER"
    Write-Host "========================================"
    Write-Host ""


    # ========================================================
    # CONNECT TO OUTLOOK
    # ========================================================

    $outlook = Connect-Outlook


    # ========================================================
    # MAPI
    # ========================================================

    Write-Host ""

    Write-Host "Accessing Outlook profile..."

    $namespace =
        $outlook.GetNamespace("MAPI")

    Write-Host "✓ MAPI profile accessed"


    # ========================================================
    # INBOX
    # ========================================================

    Write-Host ""

    Write-Host "Accessing Inbox..."

    $inbox =
        $namespace.GetDefaultFolder(6)

    Write-Host "✓ Inbox accessed"

    Write-Host ""
    Write-Host "Inbox:"
    Write-Host $inbox.FolderPath


    # ========================================================
    # CALENDAR
    # ========================================================

    Write-Host ""

    Write-Host "Accessing Calendar..."

    $calendar =
        $namespace.GetDefaultFolder(9)

    Write-Host "✓ Calendar accessed"

    Write-Host ""
    Write-Host "Calendar:"
    Write-Host $calendar.FolderPath


    # ========================================================
    # SEARCH INBOX
    #
    # We deliberately search for "Work Log ICS:" anywhere in
    # the subject.
    #
    # This catches:
    #
    # Work Log ICS: time
    # FW: Work Log ICS: time
    # RE: Work Log ICS: time
    # FW: FW: Work Log ICS: time
    #
    # We ALSO require a time.ics attachment.
    # ========================================================

    Write-Host ""

    Write-Host "Searching Inbox for Work Log ICS email..."

    $mailItems = $inbox.Items


    # Newest first

    $mailItems.Sort(
        "[ReceivedTime]",
        $true
    )


    foreach ($mail in $mailItems) {

        # 43 = Outlook MailItem

        if ($mail.Class -ne 43) {

            continue
        }


        $subject = [string]$mail.Subject


        if (
            $subject -notlike "*$SubjectSearch*"
        ) {

            continue
        }


        # ----------------------------------------------------
        # Look specifically for time.ics
        # ----------------------------------------------------

        for (
            $i = 1;
            $i -le $mail.Attachments.Count;
            $i++
        ) {

            $attachment =
                $mail.Attachments.Item($i)


            if (
                $attachment.FileName -ieq
                $AttachmentName
            ) {

                $targetMail = $mail

                $icsAttachment = $attachment

                break
            }
        }


        if ($null -ne $targetMail) {

            break
        }
    }


    # ========================================================
    # EMAIL NOT FOUND
    # ========================================================

    if ($null -eq $targetMail) {

        Write-Host ""

        Write-Host "No matching Work Log ICS email found."

        Write-Host ""
        Write-Host "Expected a message containing:"
        Write-Host "  Subject: *Work Log ICS:*"
        Write-Host "  Attachment: time.ics"

        exit 0
    }


    # ========================================================
    # EMAIL FOUND
    # ========================================================

    Write-Host ""

    Write-Host "✓ Work Log email found"

    Write-Host "  Subject:"
    Write-Host "    $($targetMail.Subject)"

    Write-Host "  Received:"
    Write-Host "    $($targetMail.ReceivedTime)"

    Write-Host "  Attachment:"
    Write-Host "    $($icsAttachment.FileName)"


    # ========================================================
    # SAVE ICS
    # ========================================================

    Write-Host ""

    Write-Host "Saving ICS attachment..."

    if (Test-Path $TempICS) {

        Remove-Item `
            $TempICS `
            -Force `
            -ErrorAction SilentlyContinue
    }


    $icsAttachment.SaveAsFile(
        $TempICS
    )


    Write-Host "✓ ICS saved"


    # ========================================================
    # PARSE ICS
    # ========================================================

    Write-Host ""

    Write-Host "Parsing ICS..."

    $events =
        Read-IcsEvents $TempICS


    Write-Host "✓ Events found: $($events.Count)"

    Write-Host ""


    # ========================================================
    # COUNTERS
    # ========================================================

    $createdCount = 0

    $updatedCount = 0

    $skippedCount = 0


    # ========================================================
    # PROCESS EACH EVENT
    # ========================================================

    foreach ($icsEvent in $events) {

        $properties =
            $icsEvent.Properties


        # ----------------------------------------------------
        # UID
        # ----------------------------------------------------

        $uid =
            $properties["UID"]


        if (
            [string]::IsNullOrWhiteSpace($uid)
        ) {

            Write-Host "Skipping event without UID."

            $skippedCount++

            continue
        }


        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------

        $summary =
            Unescape-IcsText(
                $properties["SUMMARY"]
            )


        # ----------------------------------------------------
        # DESCRIPTION
        # ----------------------------------------------------

        $description =
            Unescape-IcsText(
                $properties["DESCRIPTION"]
            )


        # ----------------------------------------------------
        # LOCATION
        # ----------------------------------------------------

        $location =
            Unescape-IcsText(
                $properties["LOCATION"]
            )


        # ----------------------------------------------------
        # DATES
        # ----------------------------------------------------

        try {

            $start =
                Parse-IcsDate(
                    $properties["DTSTART"]
                )


            $end =
                Parse-IcsDate(
                    $properties["DTEND"]
                )
        }
        catch {

            Write-Host ""
            Write-Host "Could not parse event:"
            Write-Host "  UID: $uid"
            Write-Host $_.Exception.Message

            $skippedCount++

            continue
        }


        # ====================================================
        # DISPLAY
        # ====================================================

        Write-Host "----------------------------------------"

        Write-Host "Processing event"

        Write-Host "  UID:"
        Write-Host "    $uid"

        Write-Host "  Summary:"
        Write-Host "    $summary"

        Write-Host "  Date:"
        Write-Host "    $($start.ToString('yyyy-MM-dd'))"


        if ($location) {

            Write-Host "  Location:"
            Write-Host "    $location"
        }


        # ====================================================
        # FIND EXISTING EVENT
        # ====================================================

        $existing =
            Find-ExistingWorklogEvent `
                -Calendar $calendar `
                -UID $uid


        # ====================================================
        # CREATE EVENT
        # ====================================================

        if ($null -eq $existing) {

            Write-Host "  Action: CREATE"


            $appointment =
                $calendar.Items.Add(1)


            # ------------------------------------------------
            # Basic event information
            # ------------------------------------------------

            $appointment.Subject =
                $summary


            $appointment.Start =
                $start


            $appointment.End =
                $end


            $appointment.AllDayEvent =
                $true


            $appointment.Body =
                $description


            $appointment.Location =
                $location


            # ------------------------------------------------
            # Busy status
            #
            # Outlook:
            #
            # 0 = Free
            # 1 = Tentative
            # 2 = Busy
            # 3 = Out of Office
            # 4 = Working Elsewhere
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
            # Store our stable UID
            # ------------------------------------------------

            $uidProperty =
                $appointment.UserProperties.Add(
                    "WorklogUID",
                    1,
                    $false
                )


            $uidProperty.Value =
                $uid


            # ------------------------------------------------
            # Save
            # ------------------------------------------------

            $appointment.Save()


            Write-Host "  ✓ Created"

            $createdCount++


            # ------------------------------------------------
            # Release COM objects
            # ------------------------------------------------

            if ($uidProperty) {

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $uidProperty
                ) | Out-Null
            }


            if ($appointment) {

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $appointment
                ) | Out-Null
            }
        }


        # ====================================================
        # UPDATE EXISTING EVENT
        # ====================================================

        else {

            Write-Host "  Action: UPDATE"


            $existing.Subject =
                $summary


            $existing.Start =
                $start


            $existing.End =
                $end


            $existing.AllDayEvent =
                $true


            $existing.Body =
                $description


            $existing.Location =
                $location


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
    # FINAL SUMMARY
    # ========================================================

    Write-Host ""
    Write-Host "========================================"
    Write-Host "IMPORT COMPLETE"
    Write-Host "========================================"

    Write-Host ""

    Write-Host "Events in ICS: $($events.Count)"

    Write-Host "Created:       $createdCount"

    Write-Host "Updated:       $updatedCount"

    Write-Host "Skipped:       $skippedCount"

    Write-Host ""


}
catch {

    Write-Host ""

    Write-Host "========================================"

    Write-Host "IMPORT FAILED"

    Write-Host "========================================"

    Write-Host ""

    Write-Host $_.Exception.Message

    Write-Host ""

    Write-Host $_.ScriptStackTrace

    exit 1
}


finally {

    # ========================================================
    # CLEAN UP TEMP FILE
    # ========================================================

    if (Test-Path $TempICS) {

        Remove-Item `
            $TempICS `
            -Force `
            -ErrorAction SilentlyContinue
    }


    # ========================================================
    # RELEASE COM OBJECTS
    # ========================================================

    if ($icsAttachment) {

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $icsAttachment
        ) | Out-Null
    }


    if ($targetMail) {

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $targetMail
        ) | Out-Null
    }


    if ($mailItems) {

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $mailItems
        ) | Out-Null
    }


    if ($inbox) {

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $inbox
        ) | Out-Null
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


    # ========================================================
    # FORCE COM CLEANUP
    # ========================================================

    [GC]::Collect()

    [GC]::WaitForPendingFinalizers()
}
