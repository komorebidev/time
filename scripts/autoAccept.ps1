# ============================================================
# Worklog ICS -> Outlook Calendar
#
# Reads the newest email containing:
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
#
# NEW:
#   - Only imports the newest matching email
#   - Deletes older matching Work Log ICS emails
#   - Marks successful imports
#   - Deletes imported source emails
#
# No external state folder is required.
# ============================================================


$ErrorActionPreference = "Stop"


# ============================================================
# CONFIGURATION
# ============================================================

$SubjectSearch = "Work Log ICS:"

$AttachmentName = "time.ics"


# Outlook category used before deleting processed email

$ProcessedCategory = "WorkLog Imported"


$TempICS = Join-Path `
    $env:TEMP `
    "worklog-time.ics"



# ============================================================
# CONNECT TO OUTLOOK WITH RETRIES
#
# Checks whether Outlook desktop is running.
#
# If not running:
#   - Starts Outlook.exe
#   - Waits for Outlook initialization
#   - Connects through COM
#
# Handles:
#   RPC_E_CALL_REJECTED
#   Outlook startup timing issues
# ============================================================

function Connect-Outlook {


    $maxAttempts = 12

    $delaySeconds = 5



    # --------------------------------------------------------
    # Check if Outlook desktop is already running
    # --------------------------------------------------------

    $outlookProcess =
        Get-Process `
            -Name OUTLOOK `
            -ErrorAction SilentlyContinue



    if ($null -eq $outlookProcess) {


        Write-Host ""

        Write-Host "Outlook desktop is not running."

        Write-Host "Starting Outlook..."



        Start-Process `
            "outlook.exe"



        Write-Host "Waiting for Outlook startup..."



        Start-Sleep `
            -Seconds 10
    }
    else {


        Write-Host ""

        Write-Host "Outlook desktop already running."

    }




    # --------------------------------------------------------
    # Connect COM with retries
    # --------------------------------------------------------

    for (
        $attempt = 1;
        $attempt -le $maxAttempts;
        $attempt++
    ) {


        try {


            Write-Host ""

            Write-Host "Connecting to Outlook COM..."

            Write-Host "Attempt $attempt of $maxAttempts"



            $outlook =
                New-Object `
                    -ComObject Outlook.Application



            Write-Host "✓ Outlook COM connected"



            # ------------------------------------------------
            # Verify MAPI is available
            # ------------------------------------------------

            Write-Host "Checking MAPI profile..."



            $namespace =
                $outlook.GetNamespace("MAPI")



            Write-Host "✓ MAPI available"



            return $outlook


        }
        catch {


            Write-Host ""

            Write-Host "Outlook is not ready yet."

            Write-Host $_.Exception.Message



            if (
                $attempt -lt $maxAttempts
            ) {


                Write-Host ""

                Write-Host "Waiting $delaySeconds seconds..."



                Start-Sleep `
                    -Seconds $delaySeconds

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
# 20260805
# 20260805T120000
# 20260805T120000Z
#
# ============================================================

function Parse-IcsDate {

    param(
        [string]$Value
    )


    if ([string]::IsNullOrWhiteSpace($Value)) {

        throw "ICS date value is empty."
    }


    if (
        $Value -match '^(\d{4})(\d{2})(\d{2})$'
    ) {

        return [datetime]::ParseExact(
            $Value,
            "yyyyMMdd",
            [Globalization.CultureInfo]::InvariantCulture
        )
    }


    if (
        $Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$'
    ) {

        return [datetime]::ParseExact(
            $Value,
            "yyyyMMddTHHmmssZ",
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal
        ).ToLocalTime()
    }


    if (
        $Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})$'
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


    # Normalize line endings

    $raw = $raw -replace "`r`n", "`n"

    $raw = $raw -replace "`r", "`n"


    # iCalendar line unfolding

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


        $parts = $line -split ":",2


        if ($parts.Count -ne 2) {

            continue
        }


        $propertyPart = $parts[0]

        $value = $parts[1]


        $propertyName =
            ($propertyPart -split ";")[0].ToUpper()


        $current.Properties[$propertyName] = $value
    }


    return $events
}





# ============================================================
# FIND EXISTING WORKLOG EVENT
#
# Uses Outlook custom property:
#
#     WorklogUID
#
# Prevents duplicate calendar entries.
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

                # AppointmentItem

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

                # Ignore bad calendar items
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
# FIND LATEST WORKLOG EMAIL
#
# Searches Inbox for:
#
# Subject:
#     *Work Log ICS:*
#
# Attachment:
#     time.ics
#
# Returns newest matching email.
#
# Older matching emails are returned separately so they can
# be deleted.
# ============================================================

function Find-WorklogEmails {

    param(
        $Inbox
    )


    $matches = @()


    $items = $Inbox.Items


    try {

        $items.Sort(
            "[ReceivedTime]",
            $true
        )


        foreach ($mail in $items) {


            # Outlook MailItem

            if ($mail.Class -ne 43) {

                continue
            }


            $subject =
                [string]$mail.Subject


            if (
                $subject -notlike "*$SubjectSearch*"
            ) {

                continue
            }


            $hasAttachment = $false


            for (
                $i = 1;
                $i -le $mail.Attachments.Count;
                $i++
            ) {


                $attachment =
                    $mail.Attachments.Item($i)


                if (
                    $attachment.FileName -ieq $AttachmentName
                ) {

                    $hasAttachment = $true

                    break
                }
            }


            if ($hasAttachment) {

                $matches += $mail
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


    return $matches
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

    $outlook =
        Connect-Outlook




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
    # FIND MATCHING EMAILS
    #
    # New behaviour:
    #
    # - Find all matching Work Log ICS emails
    # - Pick newest
    # - Delete older ones
    #
    # ========================================================

    Write-Host ""

    Write-Host "Searching for Work Log ICS emails..."



    $matchingEmails =
        Find-WorklogEmails `
            -Inbox $inbox



    if (
        $matchingEmails.Count -eq 0
    ) {

        Write-Host ""

        Write-Host "No Work Log ICS emails found."

        exit 0
    }




    # Newest first

    $targetMail =
        $matchingEmails[0]



    Write-Host ""

    Write-Host "Newest Work Log email selected:"


    Write-Host "Subject:"
    Write-Host $targetMail.Subject


    Write-Host "Received:"
    Write-Host $targetMail.ReceivedTime




    # ========================================================
    # DELETE OLD MATCHING EMAILS
    #
    # Keeps mailbox clean.
    #
    # Only the newest email is processed.
    #
    # ========================================================

    if (
        $matchingEmails.Count -gt 1
    ) {


        Write-Host ""

        Write-Host "Deleting older Work Log emails..."


        for (
            $i = 1;
            $i -lt $matchingEmails.Count;
            $i++
        ) {


            $oldMail =
                $matchingEmails[$i]


            Write-Host ""

            Write-Host "Deleting:"
            Write-Host $oldMail.Subject

            Write-Host $oldMail.ReceivedTime


            $oldMail.Delete()
        }


        Write-Host ""

        Write-Host "✓ Old emails deleted"
    }






    # ========================================================
    # FIND TIME.ICS ATTACHMENT
    # ========================================================

    for (
        $i = 1;
        $i -le $targetMail.Attachments.Count;
        $i++
    ) {


        $attachment =
            $targetMail.Attachments.Item($i)



        if (
            $attachment.FileName -ieq
            $AttachmentName
        ) {


            $icsAttachment =
                $attachment


            break
        }
    }




    if ($null -eq $icsAttachment) {


        throw "Selected email has no time.ics attachment."
    }




    Write-Host ""

    Write-Host "Attachment found:"
    Write-Host $icsAttachment.FileName
    # ========================================================
    # SAVE ICS FILE
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
        Read-IcsEvents `
            $TempICS


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
            Write-Host "UID: $uid"

            Write-Host $_.Exception.Message


            $skippedCount++


            continue
        }





        # ====================================================
        # DISPLAY
        # ====================================================

        Write-Host ""
        Write-Host "----------------------------------------"

        Write-Host "Processing event"


        Write-Host "UID:"
        Write-Host $uid


        Write-Host "Summary:"
        Write-Host $summary


        Write-Host "Date:"
        Write-Host $start.ToString("yyyy-MM-dd")



        if ($location) {

            Write-Host "Location:"
            Write-Host $location
        }





        # ====================================================
        # FIND EXISTING EVENT
        # ====================================================

        $existing =
            Find-ExistingWorklogEvent `
                -Calendar $calendar `
                -UID $uid
        # ====================================================
        # CREATE NEW EVENT
        # ====================================================

        if ($null -eq $existing) {


            Write-Host "Action: CREATE"



            $appointment =
                $calendar.Items.Add(1)




            # ------------------------------------------------
            # BASIC INFORMATION
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
            # BUSY STATUS
            #
            # 0 = Free
            # 1 = Tentative
            # 2 = Busy
            # 3 = Out Of Office
            #
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
            # STORE ICS UID
            #
            # Used for duplicate prevention
            #
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
            # SAVE EVENT
            # ------------------------------------------------

            $appointment.Save()



            Write-Host "✓ Created"


            $createdCount++





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


            Write-Host "Action: UPDATE"



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



            Write-Host "✓ Updated"


            $updatedCount++





            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $existing
            ) | Out-Null
        }



        Write-Host ""
    }







    # ========================================================
    # MARK EMAIL AND DELETE AFTER SUCCESSFUL IMPORT
    #
    # Category is added before deletion.
    #
    # This leaves a processing marker if Outlook delays
    # deletion.
    #
    # ========================================================

    if ($null -ne $targetMail) {


        Write-Host ""

        Write-Host "Marking email as processed..."



        if (
            [string]::IsNullOrWhiteSpace(
                $targetMail.Categories
            )
        ) {

            $targetMail.Categories =
                $ProcessedCategory

        }
        elseif (
            $targetMail.Categories -notmatch
            "(^|,\s*)$([regex]::Escape($ProcessedCategory))(,|$)"
        ) {

            $targetMail.Categories +=
                ", $ProcessedCategory"
        }



        $targetMail.Save()



        Write-Host "✓ Category added:"
        Write-Host $ProcessedCategory




        Write-Host ""

        Write-Host "Deleting processed Work Log email..."



        $targetMail.Delete()



        Write-Host "✓ Email deleted"
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
    # CLEAN TEMP ICS
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
    #
    # Prevents orphaned Outlook COM processes.
    #
    # ========================================================

    [GC]::Collect()

    [GC]::WaitForPendingFinalizers()
}



# ============================================================
# END OF SCRIPT
#
# Scheduled task behaviour:
#
# 1. Finds all emails containing:
#
#       Work Log ICS:
#
#    with:
#
#       time.ics
#
#
# 2. Keeps newest email only.
#
# 3. Deletes older matching emails.
#
# 4. Imports ICS events.
#
# 5. Creates or updates Outlook calendar items.
#
# 6. Uses WorklogUID to prevent duplicate events.
#
# 7. Marks imported email:
#
#       WorkLog Imported
#
# 8. Deletes imported source email.
#
# 9. No state folder required.
#
# ============================================================