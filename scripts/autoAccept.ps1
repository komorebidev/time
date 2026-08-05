# ============================================================
# Worklog ICS -> Outlook Calendar (Pure COM / Completely Silent)
# ============================================================

$ErrorActionPreference = "Stop"

# ============================================================
# CONFIGURATION
# ============================================================

$SubjectSearch = "Work Log ICS:"
$AttachmentName = "time.ics"
$TempICS = Join-Path $env:TEMP "worklog-time.ics"

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
    param([string]$Text)
    if ($null -eq $Text) { return "" }
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
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw "ICS date value is empty." }

    if ($Value -match '^(\d{4})(\d{2})(\d{2})$') {
        return [datetime]::ParseExact($Value, "yyyyMMdd", [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$') {
        return ([datetime]::ParseExact($Value, "yyyyMMddTHHmmssZ", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal)).ToLocalTime()
    }
    if ($Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})$') {
        return [datetime]::ParseExact($Value, "yyyyMMddTHHmmss", [Globalization.CultureInfo]::InvariantCulture)
    }
    throw "Unsupported ICS date format: $Value"
}

# ============================================================
# READ ICS EVENTS
# ============================================================

function Read-IcsEvents {
    param([string]$Path)
    $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    $raw = $raw -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n[ `t]", ""
    $lines = $raw -split "`n"
    $events = @()
    $current = $null

    foreach ($line in $lines) {
        $line = $line.TrimEnd()
        if ($line -eq "BEGIN:VEVENT") { $current = @{ Properties = @{} }; continue }
        if ($line -eq "END:VEVENT") { if ($null -ne $current) { $events += $current }; $current = $null; continue }
        if ($null -eq $current) { continue }

        $parts = $line -split ":", 2
        if ($parts.Count -ne 2) { continue }

        $propertyName = ($parts[0] -split ";")[0].ToUpper()
        $current.Properties[$propertyName] = $parts[1]
    }
    return @($events)
}

# ============================================================
# FIND EXISTING WORKLOG EVENT
# ============================================================

function Find-ExistingWorklogEvent {
    param($Calendar, [string]$UID, [string]$Summary, [datetime]$Start)
    $items = $null
    try {
        $items = $Calendar.Items
        $count = $items.Count
        for ($index = 1; $index -le $count; $index++) {
            $item = $null
            try {
                $item = $items.Item($index)
                if ($item.Class -ne 26) { continue }

                $property = $null
                try {
                    $property = $item.UserProperties.Find("WorklogUID")
                    if ($null -ne $property -and $property.Value -eq $UID) { return $item }
                }
                finally {
                    if ($property) { [Runtime.InteropServices.Marshal]::ReleaseComObject($property) | Out-Null }
                }

                if ($item.Subject -eq $Summary) {
                    $itemStart = [datetime]$item.Start
                    if ($itemStart.Date -eq $Start.Date) { return $item }
                }
            }
            catch {}
        }
    }
    finally {
        if ($items) { [Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null }
    }
    return $null
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
$skippedCount = 0
$scriptFailed = $false

# ============================================================
# MAIN
# ============================================================

try {
    Write-Host ""
    Write-Host "========================================"
    Write-Host "WORKLOG OUTLOOK IMPORTER"
    Write-Host "========================================"
    Write-Host ""

    $outlook = Connect-Outlook
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $calendar = $namespace.GetDefaultFolder(9)

    Write-Host ""
    Write-Host "Scanning Inbox items directly..."

    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    $count = $items.Count
    
    for ($index = 1; $index -le $count; $index++) {
        $mail = $items.Item($index)
        if ($mail.Class -eq 43 -and $mail.Subject -like "*$SubjectSearch*") {
            $hasTargetAttachment = $false
            foreach ($att in $mail.Attachments) {
                if ($att.FileName -ieq $AttachmentName) {
                    $hasTargetAttachment = $true
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($att) | Out-Null
                    break
                }
                [Runtime.InteropServices.Marshal]::ReleaseComObject($att) | Out-Null
            }

            if ($hasTargetAttachment) {
                $targetMail = $mail
                break
            }
        }
        [Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null
    }
    [Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null

    if ($null -eq $targetMail) {
        Write-Host "No Work Log ICS emails found."
    }
    else {
        Write-Host "Found matching email: $($targetMail.Subject)"
        
        $targetEntryID = $targetMail.EntryID
        try { $targetStoreID = $targetMail.StoreID } catch { $targetStoreID = $null }

        foreach ($att in $targetMail.Attachments) {
            if ($att.FileName -ieq $AttachmentName) {
                $icsAttachment = $att
                break
            }
            [Runtime.InteropServices.Marshal]::ReleaseComObject($att) | Out-Null
        }

        if ($null -eq $icsAttachment) {
            throw "time.ics attachment not found on target email."
        }

        if (Test-Path $TempICS) { Remove-Item $TempICS -Force -ErrorAction SilentlyContinue }
        $icsAttachment.SaveAsFile($TempICS)
        
        $events = Read-IcsEvents -Path $TempICS

        [Runtime.InteropServices.Marshal]::ReleaseComObject($icsAttachment) | Out-Null
        $icsAttachment = $null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($targetMail) | Out-Null
        $targetMail = $null

        if (Test-Path $TempICS) { Remove-Item $TempICS -Force -ErrorAction SilentlyContinue }

        foreach ($icsEvent in $events) {
            $properties = $icsEvent.Properties
            $uid = $properties["UID"]
            if ([string]::IsNullOrWhiteSpace($uid)) { $skippedCount++; continue }

            $summary = Unescape-IcsText($properties["SUMMARY"])
            $description = Unescape-IcsText($properties["DESCRIPTION"])
            $location = Unescape-IcsText($properties["LOCATION"])

            try {
                $start = Parse-IcsDate($properties["DTSTART"])
                $end = Parse-IcsDate($properties["DTEND"])
            }
            catch {
                $skippedCount++
                continue
            }

            $existing = Find-ExistingWorklogEvent -Calendar $calendar -UID $uid -Summary $summary -Start $start

            if ($null -eq $existing) {
                $appointment = $calendar.Items.Add(1)
                $appointment.Subject = $summary
                $appointment.Start = $start
                $appointment.End = $end
                $appointment.AllDayEvent = $true
                $appointment.Body = $description
                $appointment.Location = $location
                $appointment.BusyStatus = if ($summary -eq "Out of Office") { 3 } else { 2 }

                $uidProperty = $appointment.UserProperties.Add("WorklogUID", 1, $false)
                $uidProperty.Value = $uid
                $appointment.Save()

                [Runtime.InteropServices.Marshal]::ReleaseComObject($uidProperty) | Out-Null
                [Runtime.InteropServices.Marshal]::ReleaseComObject($appointment) | Out-Null
                $createdCount++
            }
            else {
                $skippedCount++
            }
        }

        # MOVE TO DELETED ITEMS VIA COM METHOD
        Write-Host ""
        Write-Host "Moving processed email to Deleted Items..."
        try {
            $mailToMove = if ($null -ne $targetStoreID) { 
                $namespace.GetItemFromID($targetEntryID, $targetStoreID) 
            } else { 
                $namespace.GetItemFromID($targetEntryID) 
            }

            if ($null -ne $mailToMove) {
                $destFolder = $namespace.GetDefaultFolder(3)
                $mailToMove.Move($destFolder)
                [Runtime.InteropServices.Marshal]::ReleaseComObject($destFolder) | Out-Null
                [Runtime.InteropServices.Marshal]::ReleaseComObject($mailToMove) | Out-Null
                Write-Host "✓ Email successfully moved."
            }
        }
        catch {
            Write-Host "Warning: Direct move encountered issue: $_"
        }
    }

}
catch {
    $scriptFailed = $true
    Write-Host ""
    Write-Host "========================================"
    Write-Host "ERROR ENCOUNTERED:" -ForegroundColor Red
    Write-Host $_ -ForegroundColor Red
    Write-Host "========================================"
}
finally {
    if ($icsAttachment) { [Runtime.InteropServices.Marshal]::ReleaseComObject($icsAttachment) | Out-Null }
    if ($targetMail) { [Runtime.InteropServices.Marshal]::ReleaseComObject($targetMail) | Out-Null }
    if ($calendar) { [Runtime.InteropServices.Marshal]::ReleaseComObject($calendar) | Out-Null }
    if ($inbox) { [Runtime.InteropServices.Marshal]::ReleaseComObject($inbox) | Out-Null }
    if ($namespace) { [Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null }
    if ($outlook) { [Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if (Test-Path $TempICS) { Remove-Item $TempICS -Force -ErrorAction SilentlyContinue }

    if ($script:StartedOutlookByScript) {
        Stop-Process -Name OUTLOOK -Force -ErrorAction SilentlyContinue
    }

    Write-Host ""
    Write-Host "========================================"
    Write-Host "FINISHED. Created: $createdCount | Skipped: $skippedCount"
    Write-Host "========================================"

    Write-Host "Closing in 5 seconds... Press any key to stay open."
    
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt 5000) {
        if ([Console]::KeyAvailable) {
            $null = [Console]::ReadKey($true)
            Write-Host "Pause requested. Press any key to exit."
            $null = [Console]::ReadKey($true)
            break
        }
        Start-Sleep -Milliseconds 100
    }
    $sw.Stop()
}