# ============================================================
# Worklog ICS -> Outlook Calendar (Pure COM / Completely Silent)
#
# Uses classic Outlook COM/MAPI without spawning an interactive process.
#
# Handles:
#   - Work Log ICS emails
#   - time.ics attachments (with forced immediate file release & deletion)
#   - Multiple VEVENT entries
#   - Advanced duplicate prevention (UID + Subject/Date fallback)
#   - Calendar updates
#   - Background COM instantiation (completely avoids UI popups)
#   - Deleting older matching emails
#   - Moving the newest processed email to the Archive folder reliably
#   - Safe process cleanup upon completion
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
# CONNECT TO OUTLOOK (PURE COM BACKGROUND INSTANTIATION)
# ============================================================

function Connect-Outlook {
    Write-Host ""
    Write-Host "Connecting to Outlook via background COM..."

    $outlook = $null
    $maxAttempts = 10

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") -as [Microsoft.Office.Interop.Outlook.Application]
            if ($null -eq $outlook) {
                throw "No active instance found."
            }
            Write-Host "✓ Connected to active Outlook instance"
            break
        }
        catch {
            try {
                $outlookType = [Type]::GetTypeFromProgID("Outlook.Application")
                $outlook = [Activator]::CreateInstance($outlookType)
                $script:StartedOutlookByScript = $true
                Write-Host "✓ Created background Outlook COM instance"
                break
            }
            catch {
                Write-Host "Attempt $attempt of $maxAttempts failed to bind Outlook COM. Retrying..."
                if ($attempt -eq $maxAttempts) {
                    throw "Failed to initialize Outlook COM interface: $_"
                }
                Start-Sleep -Seconds 3
            }
        }
    }

    return $outlook
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
# FIND EXISTING WORKLOG EVENT (ENHANCED DEDUPLICATION)
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

                # 1. Check via WorklogUID property first
                $property = $null
                try {
                    $property = $item.UserProperties.Find("WorklogUID")
                    if ($null -ne $property -and $property.Value -eq $UID) {
                        return $item
                    }
                }
                finally {
                    if ($property) {
                        [Runtime.InteropServices.Marshal]::ReleaseComObject($property) | Out-Null
                    }
                }

                # 2. Fallback check via Subject + Start Date to catch legacy/un-tagged duplicates
                if ($item.Subject -eq $Summary) {
                    $itemStart = [datetime]$item.Start
                    if ($itemStart.Date -eq $Start.Date) {
                        return $item
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
$targetEntryID = $null
$targetStoreID = $null
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

    # CONNECT OUTLOOK VIA BACKGROUND COM
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

    # CAPTURE ENTRY ID, STORE ID, AND RELEASE TARGET MAIL POINTER
    $targetEntryID = $targetMail.EntryID
    try {
        $targetStoreID = $targetMail.StoreID
    }
    catch {
        $targetStoreID = $null
    }

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

    # CLEAR ANY LINGERING TEMP ICS FILE FIRST
    if (Test-Path $TempICS) {
        Remove-Item $TempICS -Force -ErrorAction SilentlyContinue
    }

    Write-Host ""
    Write-Host "Saving ICS..."
    $icsAttachment.SaveAsFile($TempICS)
    Write-Host "✓ ICS saved"

    # PARSE EVENTS
    $events = Read-IcsEvents -Path $TempICS

    # RELEASE ATTACHMENT COM OBJECT AND TARGET MAIL IMMEDIATELY SO FILE LOCK CLEARS
    [Runtime.InteropServices.Marshal]::ReleaseComObject($icsAttachment) | Out-Null
    $icsAttachment = $null

    [Runtime.InteropServices.Marshal]::ReleaseComObject($targetMail) | Out-Null
    $targetMail = $null

    # FORCE DELETE TEMP ICS FILE IMMEDIATELY AFTER PARSING
    if (Test-Path $TempICS) {
        Remove-Item $TempICS -Force -ErrorAction SilentlyContinue
        Write-Host "✓ Temporary ICS file cleaned up"
    }

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

        # LOOK FOR EXISTING EVENT (WITH SUBJECT + DATE FALLBACK)
        $existing = Find-ExistingWorklogEvent -Calendar $calendar -UID $uid -Summary $summary -Start $start

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

    # MOVE THE NEWEST PROCESSED EMAIL TO THE ARCHIVE FOLDER (OR FALLBACK TO DELETED ITEMS)
    Write-Host ""
    Write-Host "Moving newest processed email to Archive..."
    try {
        $mailToMove = $null
        if ($null -ne $targetStoreID) {
            $mailToMove = $namespace.GetItemFromID($targetEntryID, $targetStoreID)
        }
        else {
            $mailToMove = $namespace.GetItemFromID($targetEntryID)
        }

        if ($null -ne $mailToMove) {
            # Try to get the default Archive folder (Folder Type ID 23)
            $archiveFolder = $null
            try {
                $archiveFolder = $namespace.GetDefaultFolder(23)
            }
            catch {
                # Fallback if mailbox doesn't have a formal Archive folder enabled (use Deleted Items = 3)
                $archiveFolder = $namespace.GetDefaultFolder(3)
            }

            if ($null -ne $archiveFolder) {
                $mailToMove.Move($archiveFolder)
                Write-Host "✓ Email moved successfully"
                [Runtime.InteropServices.Marshal]::ReleaseComObject($archiveFolder) | Out-Null
            }
            else {
                Write-Host "Warning: Could not resolve destination folder for archiving."
            }

            [Runtime.InteropServices.Marshal]::ReleaseComObject($mailToMove) | Out-Null
        }
        else {
            Write-Host "Warning: Could not retrieve email via Session GetItemFromID for movement."
        }
    }
    catch {
        Write-Host "Warning: Move operation failed: $_"
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

    # Final safety check to delete temp file if it somehow still exists
    if (Test-Path $TempICS) {
        Remove-Item $TempICS -Force -ErrorAction SilentlyContinue
    }

    # If the script spawned its own background COM process, terminate it cleanly so it leaves no orphaned invisible process.
    if ($script:StartedOutlookByScript) {
        Write-Host ""
        Write-Host "Closing background Outlook session..."
        Stop-Process -Name OUTLOOK -Force -ErrorAction SilentlyContinue
    }

    Write-Host ""
    Write-Host "========================================"
    Write-Host "FINISHED"
    Write-Host "Created: $createdCount | Skipped/Existing: $skippedCount"
    Write-Host "========================================"
}