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
#   - Outlook COM startup/retry handling
#
# ============================================================


$ErrorActionPreference = "Stop"





# ============================================================
# CONFIGURATION
# ============================================================

$SubjectSearch = "Work Log ICS:"

$AttachmentName = "time.ics"

$ProcessedCategory = "WorkLog Imported"



$TempICS =
    Join-Path `
        $env:TEMP `
        "worklog-time.ics"





# ============================================================
# OUTLOOK OWNERSHIP TRACKING
#
# If this script starts Outlook:
#     close it afterwards.
#
# If Outlook was already running:
#     leave it alone.
#
# ============================================================

$script:StartedOutlookByScript = $false





# ============================================================
# CONNECT TO OUTLOOK
#
# Handles:
#
#   RPC_E_CALL_REJECTED
#   Outlook startup delay
#   Existing Outlook sessions
#
# Order:
#
#   1. Attach to existing COM instance
#   2. Create Outlook COM instance
#   3. Wait for MAPI availability
#
# ============================================================

function Connect-Outlook {


    $maxAttempts = 18

    $delaySeconds = 10





    for (
        $attempt = 1;
        $attempt -le $maxAttempts;
        $attempt++
    ) {


        try {


            Write-Host ""

            Write-Host "Connecting to Outlook..."

            Write-Host "Attempt $attempt of $maxAttempts"





            # ------------------------------------------------
            # Try existing Outlook COM instance first
            # ------------------------------------------------

            try {


                $outlook =
                    [Runtime.InteropServices.Marshal]::GetActiveObject(
                        "Outlook.Application"
                    )


                Write-Host "✓ Connected to existing Outlook COM"



            }
            catch {


                Write-Host "No active Outlook COM session."



                $outlook =
                    New-Object `
                        -ComObject Outlook.Application



                $script:StartedOutlookByScript = $true



                Write-Host "✓ Started Outlook COM"

            }







            # ------------------------------------------------
            # Test MAPI
            # ------------------------------------------------

            Write-Host "Checking MAPI..."



            $namespace =
                $outlook.GetNamespace(
                    "MAPI"
                )



            # Force Outlook readiness check

            $folders =
                $namespace.Folders.Count



            Write-Host "✓ MAPI ready"





            if ($namespace) {


                [Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $namespace
                ) | Out-Null
            }





            return $outlook



        }
        catch {


            Write-Host ""

            Write-Host "Outlook COM not ready."

            Write-Host $_.Exception.Message





            # RPC_E_CALL_REJECTED

            if (
                $_.Exception.HResult -eq -2147418111
            ) {


                Write-Host ""

                Write-Host "Outlook rejected COM request."

                Write-Host "Retrying..."

            }





            if (
                $attempt -lt $maxAttempts
            ) {


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





        if (
            $line -eq "BEGIN:VEVENT"
        ) {


            $current =
                @{
                    Properties = @{}
                }


            continue
        }






        if (
            $line -eq "END:VEVENT"
        ) {


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





        if (
            $parts.Count -ne 2
        ) {


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


                        [Runtime.InteropServices.Marshal]::ReleaseComObject(
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


            [Runtime.InteropServices.Marshal]::ReleaseComObject(
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



            # Only MailItem

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


            [Runtime.InteropServices.Marshal]::ReleaseComObject(
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

    Write-Host "Loading Outlook profile..."



    $namespace =
        $outlook.GetNamespace(
            "MAPI"
        )



    Write-Host "✓ MAPI loaded"







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
    # FIND WORKLOG EMAILS
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






    # Newest email

    $targetMail =
        $matchingEmails[0]





    Write-Host ""

    Write-Host "Selected email:"

    Write-Host $targetMail.Subject

    Write-Host $targetMail.ReceivedTime








    # ========================================================
    # DELETE OLDER MATCHING EMAILS
    # ========================================================

    if (
        $matchingEmails.Count -gt 1
    ) {


        Write-Host ""

        Write-Host "Deleting older matching emails..."



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
    # SAVE ICS FILE
    # ========================================================

    if (
        Test-Path $TempICS
    ) {


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
            -Path $TempICS





    Write-Host ""

    Write-Host "Events found:"
    Write-Host $events.Count






    $createdCount = 0

    $updatedCount = 0

    $skippedCount = 0







    # ========================================================
    # PROCESS EVENTS
    # ========================================================

    foreach (
        $icsEvent in $events
    ) {



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

            Write-Host "Invalid date. Skipping."

            $skippedCount++

            continue
        }






        Write-Host ""

        Write-Host "--------------------------------"

        Write-Host "Processing:"
        Write-Host $summary





        $existing =
            Find-ExistingWorklogEvent `
                -Calendar $calendar `
                -UID $uid







        # ====================================================
        # CREATE NEW EVENT
        # ====================================================

        if (
            $null -eq $existing
        ) {


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





            [Runtime.InteropServices.Marshal]::ReleaseComObject(
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





            [Runtime.InteropServices.Marshal]::ReleaseComObject(
                $existing
            ) | Out-Null
        }



        Write-Host ""

    }







    # ========================================================
    # MARK AND DELETE PROCESSED EMAIL
    # ========================================================

    if ($targetMail) {


        Write-Host ""

        Write-Host "Marking processed email..."



        $targetMail.Categories =
            $ProcessedCategory



        $targetMail.Save()



        Write-Host "✓ Category added"




        Write-Host ""

        Write-Host "Deleting processed email..."



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

    Write-Host "Events:"
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

    if (
        Test-Path $TempICS
    ) {


        Remove-Item `
            $TempICS `
            -Force `
            -ErrorAction SilentlyContinue
    }





    # ========================================================
    # RELEASE CHILD COM OBJECTS
    #
    # Children first, Outlook last.
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


                [Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $comObject
                ) | Out-Null


            }
            catch {

                # Ignore cleanup errors

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

            Write-Host "Failed to close Outlook normally."

            Write-Host $_.Exception.Message

        }
    }






    # ========================================================
    # RELEASE OUTLOOK COM
    # ========================================================

    if ($outlook) {


        try {


            [Runtime.InteropServices.Marshal]::ReleaseComObject(
                $outlook
            ) | Out-Null


        }
        catch {

        }
    }





    # ========================================================
    # FORCE FINAL COM CLEANUP
    # ========================================================

    [GC]::Collect()

    [GC]::WaitForPendingFinalizers()
}