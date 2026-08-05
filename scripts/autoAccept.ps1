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
#   - Safe Outlook shutdown
#
# ============================================================


$ErrorActionPreference = "Stop"



# ============================================================
# CONFIGURATION
# ============================================================

$SubjectSearch =
    "Work Log ICS:"


$AttachmentName =
    "time.ics"


$ProcessedCategory =
    "WorkLog Imported"



$TempICS =
    Join-Path `
        $env:TEMP `
        "worklog-time.ics"





# ============================================================
# OUTLOOK OWNERSHIP TRACKING
#
# If this script launches Outlook:
#     close it afterwards.
#
# If Outlook was already running:
#     leave it alone.
#
# ============================================================

$script:StartedOutlookByScript =
    $false





# ============================================================
# WAIT FOR OUTLOOK COM
#
# Outlook COM can reject calls while starting.
#
# This function:
#
# 1. Starts Outlook minimized if required.
# 2. Waits for COM registration.
# 3. Attaches to Outlook.Application.
#
# ============================================================

function Connect-Outlook {


    $maxAttempts =
        30


    $delaySeconds =
        5





    # --------------------------------------------------------
    # Check if Outlook process exists
    # --------------------------------------------------------

    $outlookProcess =
        Get-Process `
            -Name OUTLOOK `
            -ErrorAction SilentlyContinue





    if ($null -eq $outlookProcess) {


        Write-Host ""

        Write-Host "Outlook is not running."

        Write-Host "Starting Outlook minimized..."



        Start-Process `
            "outlook.exe" `
            -WindowStyle Minimized



        $script:StartedOutlookByScript =
            $true



        Write-Host "Waiting for Outlook startup..."

    }
    else {


        Write-Host ""

        Write-Host "Outlook process already exists."

    }







    # --------------------------------------------------------
    # Attach to COM
    # --------------------------------------------------------

    for (
        $attempt = 1;
        $attempt -le $maxAttempts;
        $attempt++
    ) {


        try {


            Write-Host ""

            Write-Host "Connecting Outlook COM..."

            Write-Host "Attempt $attempt of $maxAttempts"





            $outlook =
                [Runtime.InteropServices.Marshal]::GetActiveObject(
                    "Outlook.Application"
                )





            Write-Host "✓ Outlook COM connected"





            # ------------------------------------------------
            # Test MAPI
            # ------------------------------------------------

            $namespace =
                $outlook.GetNamespace(
                    "MAPI"
                )



            $inbox =
                $namespace.GetDefaultFolder(
                    6
                )



            if ($null -eq $inbox) {


                throw "Inbox unavailable."

            }





            Write-Host "✓ MAPI ready"





            [Runtime.InteropServices.Marshal]::ReleaseComObject(
                $inbox
            ) | Out-Null



            [Runtime.InteropServices.Marshal]::ReleaseComObject(
                $namespace
            ) | Out-Null





            return $outlook



        }
        catch {


            Write-Host ""

            Write-Host "Outlook COM not ready."

            Write-Host $_.Exception.Message





            if (
                $_.Exception.HResult -eq -2147418111
            ) {


                Write-Host ""

                Write-Host "RPC_E_CALL_REJECTED"

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
#   20260805
#   20260805T120000
#   20260805T120000Z
#
# ============================================================

function Parse-IcsDate {

    param(
        [string]$Value
    )


    if (
        [string]::IsNullOrWhiteSpace($Value)
    ) {

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
#
# Returns all VEVENT blocks.
#
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





    # Unfold iCalendar lines

    $raw =
        $raw -replace "`n[ `t]", ""





    $lines =
        $raw -split "`n"





    $events =
        @()


    $current =
        $null





    foreach (
        $line in $lines
    ) {


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


            if (
                $null -ne $current
            ) {


                $events += $current

            }



            $current =
                $null



            continue

        }







        if (
            $null -eq $current
        ) {


            continue

        }







        $parts =
            $line -split ":",2





        if (
            $parts.Count -ne 2
        ) {


            continue

        }







        $propertyName =
            (
                $parts[0] -split ";"
            )[0].ToUpper()



        $value =
            $parts[1]






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



        # Do not use foreach directly on COM collections.
        # COM enumeration can trigger:
        #
        # "[PROPERTYGET, DISPID(0)] overload..."
        #
        # Instead access by index.

        $count =
            $items.Count



        for (
            $index = 1;
            $index -le $count;
            $index++
        ) {


            $item = $null



            try {


                $item =
                    $items.Item(
                        $index
                    )



                # AppointmentItem

                if (
                    $item.Class -ne 26
                ) {

                    continue
                }




                $property =
                    $null



                try {


                    $property =
                        $item.UserProperties.Find(
                            "WorklogUID"
                        )



                    if (
                        $null -ne $property
                    ) {


                        if (
                            $property.Value -eq $UID
                        ) {


                            return $item

                        }

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



    $matches =
        @()



    $items =
        $Inbox.Items





    try {


        $items.Sort(
            "[ReceivedTime]",
            $true
        )



        $count =
            $items.Count






        for (
            $index = 1;
            $index -le $count;
            $index++
        ) {


            $mail =
                $null



            try {


                $mail =
                    $items.Item(
                        $index
                    )





                # MailItem

                if (
                    $mail.Class -ne 43
                ) {

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






                $attachmentCount =
                    $mail.Attachments.Count





                for (
                    $a = 1;
                    $a -le $attachmentCount;
                    $a++
                ) {


                    $attachment =
                        $mail.Attachments.Item(
                            $a
                        )



                    if (
                        $attachment.FileName -ieq $AttachmentName
                    ) {


                        $hasAttachment =
                            $true



                        [Runtime.InteropServices.Marshal]::ReleaseComObject(
                            $attachment
                        ) | Out-Null



                        break

                    }



                    [Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $attachment
                    ) | Out-Null
                }







                if (
                    $hasAttachment
                ) {


                    $matches += $mail


                }
                else {


                    [Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $mail
                    ) | Out-Null

                }



            }
            catch {


                if ($mail) {


                    [Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $mail
                    ) | Out-Null

                }

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


$outlook =
    $null


$namespace =
    $null


$inbox =
    $null


$calendar =
    $null


$targetMail =
    $null


$icsAttachment =
    $null





try {


    Write-Host ""

    Write-Host "========================================"

    Write-Host "WORKLOG OUTLOOK IMPORTER"

    Write-Host "========================================"

    Write-Host ""







    # ========================================================
    # CONNECT OUTLOOK
    # ========================================================

    $outlook =
        Connect-Outlook





    # ========================================================
    # LOAD MAPI
    # ========================================================

    Write-Host ""

    Write-Host "Loading Outlook profile..."



    $namespace =
        $outlook.GetNamespace(
            "MAPI"
        )



    Write-Host "✓ MAPI loaded"







    # ========================================================
    # LOAD FOLDERS
    # ========================================================

    $inbox =
        $namespace.GetDefaultFolder(
            6
        )


    $calendar =
        $namespace.GetDefaultFolder(
            9
        )




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

    Write-Host "Searching Work Log ICS emails..."



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







    # Newest email is first

    $targetMail =
        $matchingEmails[0]





    Write-Host ""

    Write-Host "Selected email:"

    Write-Host $targetMail.Subject

    Write-Host $targetMail.ReceivedTime







    # ========================================================
    # DELETE OLD MATCHING EMAILS
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



            try {


                $oldMail.Delete()



            }
            catch {


                Write-Host "Failed deleting old email."

            }



        }


        Write-Host "✓ Old emails deleted"

    }









    # ========================================================
    # FIND TIME.ICS
    # ========================================================

    $attachmentCount =
        $targetMail.Attachments.Count




    for (
        $i = 1;
        $i -le $attachmentCount;
        $i++
    ) {


        $attachment =
            $targetMail.Attachments.Item(
                $i
            )



        if (
            $attachment.FileName -ieq $AttachmentName
        ) {


            $icsAttachment =
                $attachment



            break

        }


        [Runtime.InteropServices.Marshal]::ReleaseComObject(
            $attachment
        ) | Out-Null

    }







    if (
        $null -eq $icsAttachment
    ) {


        throw "time.ics attachment not found."

    }







    # ========================================================
    # SAVE ICS
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
    # PARSE EVENTS
    # ========================================================

    $events =
        Read-IcsEvents `
            -Path $TempICS




    Write-Host ""

    Write-Host "Events found:"
    Write-Host $events.Count






    $createdCount =
        0


    $updatedCount =
        0


    $skippedCount =
        0

    # ========================================================
    # PROCESS EVENTS
    # ========================================================

    foreach (
        $icsEvent in $events
    ) {


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
        # EVENT DETAILS
        # ----------------------------------------------------

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

            Write-Host "Invalid date. Skipping event."

            $skippedCount++

            continue

        }







        Write-Host ""

        Write-Host "--------------------------------"

        Write-Host "Processing:"
        Write-Host $summary








        # ----------------------------------------------------
        # CHECK EXISTING EVENT
        # ----------------------------------------------------

        $existing =
            Find-ExistingWorklogEvent `
                -Calendar $calendar `
                -UID $uid







        # ====================================================
        # CREATE EVENT
        # ====================================================

        if (
            $null -eq $existing
        ) {


            Write-Host "Action: CREATE"




            $appointment =
                $calendar.Items.Add(
                    1
                )




            try {


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


                    $appointment.BusyStatus =
                        3

                }
                else {


                    $appointment.BusyStatus =
                        2

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




                if ($uidProperty) {


                    [Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $uidProperty
                    ) | Out-Null

                }



            }
            finally {


                if ($appointment) {


                    [Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $appointment
                    ) | Out-Null

                }

            }

        }







        # ====================================================
        # UPDATE EVENT
        # ====================================================

        else {


            Write-Host "Action: UPDATE"




            try {


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


                    $existing.BusyStatus =
                        3

                }
                else {


                    $existing.BusyStatus =
                        2

                }





                $existing.Save()



                Write-Host "✓ Updated"



                $updatedCount++


            }
            finally {


                if ($existing) {


                    [Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $existing
                    ) | Out-Null

                }

            }


        }


    }







    # ========================================================
    # MARK AND DELETE SOURCE EMAIL
    # ========================================================

    if (
        $targetMail
    ) {


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
    # SUMMARY
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
    # RELEASE COM OBJECTS
    #
    # Release children first.
    # Outlook last.
    #
    # ========================================================


    foreach (
        $object in @(
            $icsAttachment,
            $targetMail,
            $calendar,
            $inbox,
            $namespace
        )
    ) {


        if (
            $null -ne $object
        ) {


            try {


                [Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $object
                ) | Out-Null


            }
            catch {

                # Ignore cleanup failures

            }

        }

    }







    # ========================================================
    # CLOSE OUTLOOK IF SCRIPT STARTED IT
    #
    # If the user already had Outlook open:
    #     leave it running.
    #
    # ========================================================

    if (
        $script:StartedOutlookByScript -and
        $null -ne $outlook
    ) {


        Write-Host ""

        Write-Host "Closing Outlook..."



        try {


            $outlook.Quit()



            Start-Sleep `
                -Seconds 5



            Write-Host "✓ Outlook closed"


        }
        catch {


            Write-Host ""

            Write-Host "Normal Outlook shutdown failed."

            Write-Host $_.Exception.Message



            # Fallback:
            # Only kill Outlook if we started it.

            try {


                Get-Process `
                    -Name OUTLOOK `
                    -ErrorAction SilentlyContinue |
                    Stop-Process `
                        -Force



                Write-Host "✓ Outlook process terminated"


            }
            catch {

                Write-Host "Unable to terminate Outlook."

            }

        }

    }







    # ========================================================
    # RELEASE OUTLOOK COM
    # ========================================================

    if (
        $null -ne $outlook
    ) {


        try {


            [Runtime.InteropServices.Marshal]::ReleaseComObject(
                $outlook
            ) | Out-Null


        }
        catch {

            # Ignore cleanup errors

        }

    }








    # ========================================================
    # FORCE COM GARBAGE COLLECTION
    # ========================================================

    [GC]::Collect()

    [GC]::WaitForPendingFinalizers()

}




# ============================================================
# END
#
# Features:
#
# ✓ Starts Outlook minimized when required
# ✓ Waits for Outlook COM registration
# ✓ Handles RPC_E_CALL_REJECTED
# ✓ Avoids COM collection enumeration crashes
# ✓ Imports newest Work Log ICS email
# ✓ Removes older matching emails
# ✓ Creates/updates calendar events
# ✓ Prevents duplicates with WorklogUID
# ✓ Deletes processed emails
# ✓ Closes Outlook only when started by script
#
# ============================================================