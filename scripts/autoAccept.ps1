# ============================================================
# Worklog ICS -> Outlook Calendar
#
# Reads newest email containing:
#
#     Work Log ICS:
#
# with a time.ics attachment.
#
# Uses classic Outlook COM/MAPI.
#
# Handles:
#   - Multiple VEVENTs
#   - All-day events
#   - Descriptions
#   - Locations
#   - Out Of Office events
#   - Duplicate prevention
#   - Updating existing events
#   - Outlook startup timing problems
#
# NEW:
#   - Outlook started only through COM
#   - Outlook closed cleanly after import
#   - Existing Outlook sessions are not closed
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
# OUTLOOK SESSION TRACKING
#
# True:
#   Script created Outlook session.
#
# False:
#   Outlook was already running.
#
# Only close Outlook when this script owns it.
#
# ============================================================

$script:StartedOutlookByScript = $false






# ============================================================
# CONNECT TO OUTLOOK WITH RETRIES
#
# Does NOT use Start-Process.
#
# Outlook COM launches Outlook when required.
#
# This avoids:
#   - Visible startup windows
#   - Startup race conditions
#   - WindowStyle issues
#
# ============================================================

function Connect-Outlook {


    $maxAttempts = 12

    $delaySeconds = 5




    # --------------------------------------------------------
    # Check existing Outlook process
    # --------------------------------------------------------

    $outlookProcess =
        Get-Process `
            -Name OUTLOOK `
            -ErrorAction SilentlyContinue




    if ($null -eq $outlookProcess) {


        Write-Host ""

        Write-Host "Outlook desktop is not running."

        Write-Host "Starting through COM..."



        $script:StartedOutlookByScript = $true

    }
    else {


        Write-Host ""

        Write-Host "Outlook desktop already running."

    }






    # --------------------------------------------------------
    # Connect COM with retry
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




            Write-Host "Checking MAPI profile..."



            $namespace =
                $outlook.GetNamespace("MAPI")




            Write-Host "✓ MAPI available"




            # Release temporary namespace object

            if ($namespace) {


                [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $namespace
                ) | Out-Null
            }




            return $outlook



        }
        catch {


            Write-Host ""

            Write-Host "Outlook not ready."

            Write-Host $_.Exception.Message



            if (
                $attempt -lt $maxAttempts
            ) {


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



    $Text =
        $Text -replace '\\n', "`r`n"



    $Text =
        $Text -replace '\\N', "`r`n"



    $Text =
        $Text -replace '\\,', ","



    $Text =
        $Text -replace '\\;', ";"



    $Text =
        $Text -replace '\\\\', "\"


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

    $raw =
        $raw -replace "`r`n", "`n"

    $raw =
        $raw -replace "`r", "`n"



    # iCalendar line unfolding

    $raw =
        $raw -replace "`n[ `t]", ""



    $lines =
        $raw -split "`n"



    $events = @()

    $current = $null



    foreach ($line in $lines) {


        $line =
            $line.TrimEnd()



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




        $propertyPart =
            $parts[0]


        $value =
            $parts[1]




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
# Uses custom Outlook property:
#
#     WorklogUID
#
# Prevents duplicates.
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


                # Ignore invalid calendar items

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
# FIND WORKLOG EMAILS
#
# Searches Inbox for:
#
# Subject:
#     Work Log ICS:
#
# Attachment:
#     time.ics
#
# Returns newest first.
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



            # MailItem only

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





            $hasAttachment =
                $false




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


                    $hasAttachment =
                        $true


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
    # ACCESS MAPI
    # ========================================================

    Write-Host ""

    Write-Host "Accessing Outlook profile..."



    $namespace =
        $outlook.GetNamespace("MAPI")



    Write-Host "✓ MAPI accessed"






    # ========================================================
    # ACCESS FOLDERS
    # ========================================================

    $inbox =
        $namespace.GetDefaultFolder(6)


    $calendar =
        $namespace.GetDefaultFolder(9)



    Write-Host ""

    Write-Host "Inbox:"
    Write-Host $inbox.FolderPath


    Write-Host ""

    Write-Host "Calendar:"
    Write-Host $calendar.FolderPath





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

        Write-Host "No matching emails found."

        exit 0
    }





    # Newest email

    $targetMail =
        $matchingEmails[0]




    Write-Host ""

    Write-Host "Selected email:"

    Write-Host $targetMail.Subject

    Write-Host $targetMail.ReceivedTime





    # ========================================================
    # REMOVE OLD EMAILS
    # ========================================================

    if ($matchingEmails.Count -gt 1) {


        Write-Host ""

        Write-Host "Removing older Work Log emails..."



        for (
            $i = 1;
            $i -lt $matchingEmails.Count;
            $i++
        ) {


            $matchingEmails[$i].Delete()
        }


        Write-Host "✓ Old emails removed"
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


        throw "time.ics attachment not found."
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

    Write-Host "Saving ICS..."



    $icsAttachment.SaveAsFile(
        $TempICS
    )



    Write-Host "✓ ICS saved"






    # ========================================================
    # READ EVENTS
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


            Write-Host "Skipping event without UID."

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


            Write-Host ""

            Write-Host "Skipping invalid date."

            $skippedCount++

            continue
        }





        Write-Host ""

        Write-Host "--------------------------------"

        Write-Host "Processing:"
        Write-Host $summary

        Write-Host "UID:"
        Write-Host $uid





        # ====================================================
        # CHECK EXISTING EVENT
        # ====================================================

        $existing =
            Find-ExistingWorklogEvent `
                -Calendar $calendar `
                -UID $uid






        # ====================================================
        # CREATE EVENT
        # ====================================================

        if ($null -eq $existing) {


            Write-Host "Action: CREATE"



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



            Write-Host "✓ Created"



            $createdCount++





            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $appointment
            ) | Out-Null
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
    # MARK AND DELETE IMPORTED EMAIL
    # ========================================================

    if ($targetMail) {


        Write-Host ""

        Write-Host "Marking imported email..."



        $targetMail.Categories =
            $ProcessedCategory



        $targetMail.Save()



        Write-Host "✓ Category added"





        Write-Host ""

        Write-Host "Deleting imported email..."



        $targetMail.Delete()



        Write-Host "✓ Email deleted"

    }







    # ========================================================
    # SUMMARY
    # ========================================================

    Write-Host ""

    Write-Host "========================================"

    Write-Host "IMPORT COMPLETE"

    Write-Host "========================================"

    Write-Host ""

    Write-Host "Events found:"
    Write-Host $events.Count

    Write-Host ""

    Write-Host "Created:"
    Write-Host $createdCount

    Write-Host ""

    Write-Host "Updated:"
    Write-Host $updatedCount

    Write-Host ""

    Write-Host "Skipped:"
    Write-Host $skippedCount





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
    # REMOVE TEMP ICS
    # ========================================================

    if (Test-Path $TempICS) {


        Remove-Item `
            $TempICS `
            -Force `
            -ErrorAction SilentlyContinue
    }





    # ========================================================
    # RELEASE CHILD COM OBJECTS
    #
    # Release children before Outlook.
    #
    # ========================================================

    foreach (
        $comObject in @(
            $icsAttachment,
            $targetMail,
            $mailItems,
            $inbox,
            $calendar,
            $namespace
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





    # ========================================================
    # CLOSE OUTLOOK ONLY IF SCRIPT STARTED IT
    # ========================================================

    if (
        $script:StartedOutlookByScript -and
        $outlook
    ) {


        Write-Host ""

        Write-Host "Closing Outlook..."



        try {


            $outlook.Quit()



            Start-Sleep `
                -Seconds 5



            Write-Host "✓ Outlook closed cleanly"

        }
        catch {


            Write-Host ""

            Write-Host "Outlook did not close cleanly."

            Write-Host $_.Exception.Message

        }
    }





    # ========================================================
    # RELEASE OUTLOOK COM OBJECT LAST
    # ========================================================

    if ($outlook) {


        try {


            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $outlook
            ) | Out-Null


        }
        catch {

            # Ignore

        }
    }





    # ========================================================
    # FINAL COM CLEANUP
    # ========================================================

    [GC]::Collect()

    [GC]::WaitForPendingFinalizers()
}