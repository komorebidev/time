# ============================================================
# Worklog ICS -> Outlook Calendar (Pure COM / Completely Silent)
#
# Fixed version:
# - Keeps processed email COM reference alive until Move()
# - Reliable Deleted Items move
# - Safe COM cleanup
# - Handles multiple VEVENT entries
# - UID + Subject/Date duplicate prevention
#
# ============================================================

$ErrorActionPreference = "Stop"


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
# CONNECT TO OUTLOOK
# ============================================================

function Connect-Outlook {

    Write-Host ""
    Write-Host "Connecting to Outlook via COM..."

    $outlook = $null

    for ($attempt = 1; $attempt -le 10; $attempt++) {

        try {

            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject(
                "Outlook.Application"
            )

            if ($null -ne $outlook) {

                Write-Host "✓ Connected to existing Outlook"
                return $outlook

            }

        }
        catch {

        }


        try {

            $outlookType = [Type]::GetTypeFromProgID(
                "Outlook.Application"
            )

            $outlook = [Activator]::CreateInstance(
                $outlookType
            )

            $script:StartedOutlookByScript = $true

            Write-Host "✓ Created Outlook COM instance"

            return $outlook

        }
        catch {

            Write-Host "Waiting for Outlook COM..."

            Start-Sleep -Seconds 3

        }

    }


    throw "Unable to initialize Outlook COM."

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

        throw "Empty ICS date."

    }



    if ($Value -match '^(\d{8})$') {

        return [datetime]::ParseExact(
            $Value,
            "yyyyMMdd",
            [Globalization.CultureInfo]::InvariantCulture
        )

    }



    if ($Value -match '^(\d{8}T\d{6})Z$') {

        return (
            [datetime]::ParseExact(
                $Value,
                "yyyyMMddTHHmmssZ",
                [Globalization.CultureInfo]::InvariantCulture,
                [DateTimeStyles]::AssumeUniversal
            )
        ).ToLocalTime()

    }



    if ($Value -match '^(\d{8}T\d{6})$') {

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


    # ICS unfolding

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



        $propertyName = (
            $parts[0] -split ";"
        )[0].ToUpper()



        $current.Properties[$propertyName] = $parts[1]


    }



    return @($events)

}

# ============================================================
# FIND EXISTING WORKLOG EVENT
# ============================================================

function Find-ExistingWorklogEvent {

    param(
        $Calendar,
        [string]$UID,
        [string]$Summary,
        [datetime]$Start
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



                # UID lookup

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
                finally {


                    if ($property) {

                        [Runtime.InteropServices.Marshal]::ReleaseComObject(
                            $property
                        ) | Out-Null

                    }

                }



                # Subject/date fallback

                if ($item.Subject -eq $Summary) {


                    $itemStart = [datetime]$item.Start


                    if ($itemStart.Date -eq $Start.Date) {

                        return $item

                    }

                }


            }
            catch {

                continue

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
# ============================================================

function Find-WorklogEmails {

    param(
        $Inbox
    )


    $matches = @()

    $items = $null



    try {


        $items = $Inbox.Items


        $items.Sort(
            "[ReceivedTime]",
            $true
        )


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



                $foundAttachment = $false



                $attachmentCount = $mail.Attachments.Count



                for ($a = 1; $a -le $attachmentCount; $a++) {


                    $attachment = $null



                    try {


                        $attachment = $mail.Attachments.Item($a)



                        if ($attachment.FileName -ieq $AttachmentName) {


                            $foundAttachment = $true

                            break

                        }


                    }
                    finally {


                        if ($attachment) {

                            [Runtime.InteropServices.Marshal]::ReleaseComObject(
                                $attachment
                            ) | Out-Null

                        }

                    }


                }



                if ($foundAttachment) {


                    # Keep COM object alive for processing

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



    return @($matches)

}





# ============================================================
# CREATE CALENDAR EVENT
# ============================================================

function Create-WorklogCalendarEvent {


    param(
        $Calendar,
        $Properties
    )



    $uid = $Properties["UID"]


    $summary = Unescape-IcsText(
        $Properties["SUMMARY"]
    )


    $description = Unescape-IcsText(
        $Properties["DESCRIPTION"]
    )


    $location = Unescape-IcsText(
        $Properties["LOCATION"]
    )



    $start = Parse-IcsDate(
        $Properties["DTSTART"]
    )


    $end = Parse-IcsDate(
        $Properties["DTEND"]
    )



    $existing = Find-ExistingWorklogEvent `
        -Calendar $Calendar `
        -UID $uid `
        -Summary $summary `
        -Start $start



    if ($null -ne $existing) {


        Write-Host "Already exists:"
        Write-Host $summary

        return "SKIPPED"


    }



    $appointment = $null
    $uidProperty = $null



    try {


        $appointment = $Calendar.Items.Add(1)



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



        $uidProperty = $appointment.UserProperties.Add(
            "WorklogUID",
            1,
            $false
        )


        $uidProperty.Value = $uid



        $appointment.Save()



        Write-Host "Created:"
        Write-Host $summary



        return "CREATED"


    }
    finally {


        if ($uidProperty) {

            [Runtime.InteropServices.Marshal]::ReleaseComObject(
                $uidProperty
            ) | Out-Null

        }


        if ($appointment) {

            [Runtime.InteropServices.Marshal]::ReleaseComObject(
                $appointment
            ) | Out-Null

        }


    }


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
$skippedCount = 0


# ============================================================
# MAIN
# ============================================================

try {


    Write-Host ""
    Write-Host "========================================"
    Write-Host "WORKLOG OUTLOOK IMPORTER"
    Write-Host "========================================"



    # CONNECT

    $outlook = Connect-Outlook



    # LOAD MAPI

    $namespace = $outlook.GetNamespace("MAPI")



    $inbox = $namespace.GetDefaultFolder(6)

    $calendar = $namespace.GetDefaultFolder(9)



    Write-Host ""
    Write-Host "Searching Work Log ICS emails..."



    $matchingEmails = @(Find-WorklogEmails -Inbox $inbox)



    if ($matchingEmails.Count -eq 0) {


        Write-Host "No Work Log ICS emails found."

        exit 0

    }



    # newest email

    $targetMail = $matchingEmails[0]



    Write-Host ""
    Write-Host "Selected:"
    Write-Host $targetMail.Subject
    Write-Host $targetMail.ReceivedTime





    # ========================================================
    # DELETE OLDER MATCHING EMAILS
    # ========================================================

    if ($matchingEmails.Count -gt 1) {


        Write-Host ""
        Write-Host "Deleting older emails..."



        for ($i = 1; $i -lt $matchingEmails.Count; $i++) {


            $oldMail = $matchingEmails[$i]



            try {


                $oldMail.Delete()

                Write-Host "✓ Deleted old email"


            }
            catch {


                Write-Host "Could not delete old email"


            }
            finally {


                if ($oldMail) {


                    [Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $oldMail
                    ) | Out-Null


                }


            }


        }


    }





    # ========================================================
    # FIND ICS ATTACHMENT
    # ========================================================

    $attachmentCount = $targetMail.Attachments.Count



    for ($i = 1; $i -le $attachmentCount; $i++) {


        $attachment = $targetMail.Attachments.Item($i)



        if ($attachment.FileName -ieq $AttachmentName) {


            $icsAttachment = $attachment

            break


        }


        [Runtime.InteropServices.Marshal]::ReleaseComObject(
            $attachment
        ) | Out-Null


    }



    if ($null -eq $icsAttachment) {


        throw "time.ics not found."


    }





    # ========================================================
    # SAVE ICS TEMP FILE
    # ========================================================

    if (Test-Path $TempICS) {

        Remove-Item $TempICS -Force

    }



    Write-Host ""
    Write-Host "Saving ICS..."



    $icsAttachment.SaveAsFile(
        $TempICS
    )



    Write-Host "✓ Saved"





    # ========================================================
    # READ EVENTS
    # ========================================================

    $events = Read-IcsEvents(
        $TempICS
    )



    Write-Host ""
    Write-Host "Events:"
    Write-Host $events.Count





    # Release attachment only
    # IMPORTANT:
    # Keep targetMail alive until Move()

    [Runtime.InteropServices.Marshal]::ReleaseComObject(
        $icsAttachment
    ) | Out-Null


    $icsAttachment = $null



    if (Test-Path $TempICS) {


        Remove-Item $TempICS -Force -ErrorAction SilentlyContinue


    }





    # ========================================================
    # IMPORT EVENTS
    # ========================================================

    foreach ($event in $events) {


        try {


            $result = Create-WorklogCalendarEvent `
                -Calendar $calendar `
                -Properties $event.Properties



            if ($result -eq "CREATED") {

                $createdCount++

            }
            else {

                $skippedCount++

            }



        }
        catch {


            Write-Host "Event failed:"
            Write-Host $_


            $skippedCount++


        }


    }





    # ========================================================
    # MOVE PROCESSED EMAIL
    # ========================================================

    Write-Host ""
    Write-Host "Moving processed email to Deleted Items..."



    $deletedFolder = $null



    try {


        if ($null -eq $targetMail) {

            throw "Email reference lost."

        }



        $deletedFolder = $namespace.GetDefaultFolder(3)



        $targetMail.Move(
            $deletedFolder
        )



        Write-Host "✓ Email moved to Deleted Items"


    }
    catch {


        Write-Host "Move failed:"
        Write-Host $_


    }
    finally {


        if ($deletedFolder) {


            [Runtime.InteropServices.Marshal]::ReleaseComObject(
                $deletedFolder
            ) | Out-Null


        }


    }



}
finally {


    # ========================================================
    # CLEANUP
    # ========================================================


    if ($icsAttachment) {

        [Runtime.InteropServices.Marshal]::ReleaseComObject(
            $icsAttachment
        ) | Out-Null

    }



    if ($targetMail) {

        [Runtime.InteropServices.Marshal]::ReleaseComObject(
            $targetMail
        ) | Out-Null

    }



    if ($calendar) {

        [Runtime.InteropServices.Marshal]::ReleaseComObject(
            $calendar
        ) | Out-Null

    }



    if ($inbox) {

        [Runtime.InteropServices.Marshal]::ReleaseComObject(
            $inbox
        ) | Out-Null

    }



    if ($namespace) {

        [Runtime.InteropServices.Marshal]::ReleaseComObject(
            $namespace
        ) | Out-Null

    }



    if ($outlook) {

        [Runtime.InteropServices.Marshal]::ReleaseComObject(
            $outlook
        ) | Out-Null

    }



    [GC]::Collect()

    [GC]::WaitForPendingFinalizers()



    if (Test-Path $TempICS) {

        Remove-Item $TempICS -Force -ErrorAction SilentlyContinue

    }



    if ($script:StartedOutlookByScript) {


        Write-Host ""
        Write-Host "Closing background Outlook..."

        Stop-Process `
            -Name OUTLOOK `
            -Force `
            -ErrorAction SilentlyContinue


    }



    Write-Host ""
    Write-Host "========================================"
    Write-Host "FINISHED"
    Write-Host "Created: $createdCount"
    Write-Host "Skipped: $skippedCount"
    Write-Host "========================================"


}