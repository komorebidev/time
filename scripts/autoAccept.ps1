# ============================================================
# Worklog ICS -> Outlook Calendar
#
# Reads the newest email containing:
#
#     Work Log ICS:
#
# and a time.ics attachment.
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
#   - Automatically closes Outlook if started by script
#
# ============================================================


$ErrorActionPreference = "Stop"


# ============================================================
# CONFIGURATION
# ============================================================

$SubjectSearch = "Work Log ICS:"

$AttachmentName = "time.ics"

$ProcessedCategory = "WorkLog Imported"


$TempICS = Join-Path `
    $env:TEMP `
    "worklog-time.ics"



# ============================================================
# OUTLOOK OWNERSHIP FLAG
#
# Tracks whether this script launched Outlook.
#
# If true:
#   Close Outlook when import finishes.
#
# If false:
#   Leave existing Outlook session alone.
#
# ============================================================

$script:StartedOutlookByScript = $false




# ============================================================
# CONNECT TO OUTLOOK WITH RETRIES
#
# Checks whether Outlook desktop is running.
#
# If not running:
#   - Starts Outlook minimized
#   - Waits for initialization
#   - Connects through COM
#
# Handles:
#   RPC_E_CALL_REJECTED
#   Outlook startup timing issues
#
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
            "outlook.exe" `
            -WindowStyle Minimized



        # Remember that this script owns Outlook

        $script:StartedOutlookByScript = $true



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



            # Return Outlook object

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


    $raw =
        [System.IO.File]::ReadAllText(
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



        $parts =
            $line -split ":",2



        if ($parts.Count -ne 2) {

            continue
        }



        $propertyPart = $parts[0]

        $value = $parts[1]



        $propertyName =
            ($propertyPart -split ";")[0].ToUpper()



        $current.Properties[$propertyName] =
            $value
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
#
# ============================================================

function Find-ExistingWorklogEvent {

    param(
        $Calendar,
        [string]$UID
    )


    $items = $null


    try {


        $items =
            $Calendar.Items



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

                # Ignore invalid calendar entries

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
# ============================================================

function Find-WorklogEmails {

    param(
        $Inbox
    )


    $matches = @()


    $items =
        $Inbox.Items



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





    # ========================================================
    # CALENDAR
    # ========================================================

    Write-Host ""

    Write-Host "Accessing Calendar..."


    $calendar =
        $namespace.GetDefaultFolder(9)


    Write-Host "✓ Calendar accessed"





    # ========================================================
    # FIND EMAILS
    # ========================================================

    Write-Host ""

    Write-Host "Searching for Work Log ICS emails..."



    $matchingEmails =
        Find-WorklogEmails `
            -Inbox $inbox



    if ($matchingEmails.Count -eq 0) {


        Write-Host ""

        Write-Host "No Work Log ICS emails found."


        exit 0
    }





    $targetMail =
        $matchingEmails[0]



    Write-Host ""

    Write-Host "Selected email:"

    Write-Host $targetMail.Subject

    Write-Host $targetMail.ReceivedTime





    # ========================================================
    # DELETE OLD EMAILS
    # ========================================================

    if ($matchingEmails.Count -gt 1) {


        Write-Host ""

        Write-Host "Deleting older Work Log emails..."



        for (
            $i = 1;
            $i -lt $matchingEmails.Count;
            $i++
        ) {


            $matchingEmails[$i].Delete()
        }


        Write-Host "✓ Old emails deleted"
    }





    # ========================================================
    # FIND ICS ATTACHMENT
    # ========================================================

    for (
        $i = 1;
        $i -le $targetMail.Attachments.Count;
        $i++
    ) {


        $attachment =
            $targetMail.Attachments.Item($i)



        if (
            $attachment.FileName -ieq $AttachmentName
        ) {


            $icsAttachment =
                $attachment


            break
        }
    }




    if ($null -eq $icsAttachment) {


        throw "Selected email has no time.ics attachment."
    }





    # ========================================================
    # SAVE ICS
    # ========================================================

    if (Test-Path $TempICS) {


        Remove-Item `
            $TempICS `
            -Force `
            -ErrorAction SilentlyContinue
    }



    Write-Host ""

    Write-Host "Saving ICS attachment..."


    $icsAttachment.SaveAsFile(
        $TempICS
    )


    Write-Host "✓ ICS saved"





    # ========================================================
    # PARSE EVENTS
    # ========================================================

    $events =
        Read-IcsEvents `
            $TempICS



    Write-Host ""

    Write-Host "Events found: $($events.Count)"





    $createdCount = 0

    $updatedCount = 0

    $skippedCount = 0





    # ========================================================
    # PROCESS EVENTS
    # ========================================================

    foreach ($icsEvent in $events) {


        $properties =
            $icsEvent.Properties



        $uid =
            $properties["UID"]



        if (
            [string]::IsNullOrWhiteSpace($uid)
        ) {


            $skippedCount++

            continue
        }



        $summary =
            Unescape-IcsText(
                $properties["SUMMARY"]
            )



        $description =
            Unescape-IcsText(
                $properties["DESCRIPTION"]
            )



        $location =
            Unescape-IcsText(
                $properties["LOCATION"]
            )



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


            $skippedCount++

            continue
        }





        $existing =
            Find-ExistingWorklogEvent `
                -Calendar $calendar `
                -UID $uid





        if ($null -eq $existing) {


            Write-Host "Creating: $summary"



            $appointment =
                $calendar.Items.Add(1)



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



            if (
                $summary -eq "Out of Office"
            ) {

                $appointment.BusyStatus = 3
            }
            else {

                $appointment.BusyStatus = 2
            }



            $uidProperty =
                $appointment.UserProperties.Add(
                    "WorklogUID",
                    1,
                    $false
                )



            $uidProperty.Value =
                $uid



            $appointment.Save()



            $createdCount++



            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $appointment
            ) | Out-Null
        }
        else {


            Write-Host "Updating: $summary"



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



            $existing.Save()



            $updatedCount++



            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $existing
            ) | Out-Null
        }
    }





    # ========================================================
    # DELETE PROCESSED EMAIL
    # ========================================================

    if ($targetMail) {


        $targetMail.Categories =
            $ProcessedCategory


        $targetMail.Save()


        $targetMail.Delete()


        Write-Host "✓ Source email deleted"
    }





    Write-Host ""

    Write-Host "========================================"

    Write-Host "IMPORT COMPLETE"

    Write-Host "========================================"

    Write-Host ""

    Write-Host "Created: $createdCount"

    Write-Host "Updated: $updatedCount"

    Write-Host "Skipped: $skippedCount"


}
catch {


    Write-Host ""

    Write-Host "IMPORT FAILED"

    Write-Host $_.Exception.Message


    exit 1
}



finally {


    # ========================================================
    # REMOVE TEMP FILE
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

    foreach (
        $comObject in @(
            $icsAttachment,
            $targetMail,
            $mailItems,
            $inbox,
            $calendar,
            $namespace,
            $outlook
        )
    ) {


        if ($comObject) {


            try {


                [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $comObject
                ) | Out-Null


            }
            catch {

                # Ignore cleanup failures

            }
        }
    }





    # Force COM cleanup

    [GC]::Collect()

    [GC]::WaitForPendingFinalizers()





    # ========================================================
    # CLOSE OUTLOOK IF THIS SCRIPT STARTED IT
    # ========================================================

    if ($script:StartedOutlookByScript) {


        Write-Host ""

        Write-Host "Closing Outlook started by importer..."



        try {


            Get-Process `
                -Name OUTLOOK `
                -ErrorAction SilentlyContinue |
                Stop-Process `
                    -Force



            Write-Host "✓ Outlook closed"

        }
        catch {


            Write-Host "Could not close Outlook."

        }
    }
}