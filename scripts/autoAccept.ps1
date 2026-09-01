# ============================================================
# Worklog ICS -> Outlook Calendar (OneDrive Polling / Pure COM)
# ============================================================

$ErrorActionPreference = "Stop"

# Force UTF-8 encoding for console output
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ============================================================
# CONFIGURATION (Dynamically locates OneDrive to avoid dash issues)
# ============================================================

$BaseOneDrive = Get-ChildItem "$env:USERPROFILE\OneDrive*" | Where-Object { $_.Name -like "*エイラシステム*" } | Select-Object -ExpandProperty FullName

if (-not $BaseOneDrive) {
    # Fallback if dynamic search fails
    $BaseOneDrive = "$env:USERPROFILE\OneDrive - エイラシステム株式会社"
}

$LocalFolder = Join-Path $BaseOneDrive "ics"
$Filter = "time.ics"
$TargetIcsFile = Join-Path $LocalFolder $Filter

if (-not (Test-Path $LocalFolder)) {
    New-Item -ItemType Directory -Path $LocalFolder -Force | Out-Null
}

$script:StartedOutlookByScript = $false

# ============================================================
# CONNECT TO OUTLOOK (PURE COM BACKGROUND INSTANTIATION)
# ============================================================

function Connect-Outlook {
    Write-Output "Connecting to Outlook via background COM..."

    $outlook = $null
    $maxAttempts = 10

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") -as [Microsoft.Office.Interop.Outlook.Application]
            if ($null -eq $outlook) {
                throw "No active instance found."
            }
            Write-Output "[OK] Connected to active Outlook instance"
            break
        }
        catch {
            try {
                $outlookType = [Type]::GetTypeFromProgID("Outlook.Application")
                $outlook = [Activator]::CreateInstance($outlookType)
                $script:StartedOutlookByScript = $true
                Write-Output "[OK] Created background Outlook COM instance"
                break
            }
            catch {
                Write-Output "Attempt $attempt of $maxAttempts failed to bind Outlook COM. Retrying..."
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
        return [datetime]::ParseExact($Value, "yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$') {
        $dt = [datetime]::ParseExact($Value, "yyyyMMddTHHmmssZ", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal)
        return $dt.ToLocalTime()
    }
    if ($Value -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})$') {
        return [datetime]::ParseExact($Value, "yyyyMMddTHHmmss", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    throw "Unsupported ICS date format: $Value"
}

# ============================================================
# READ ICS EVENTS
# ============================================================

function Read-IcsEvents {
    param([string]$Path)
    
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $raw = ""

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $raw = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    } else {
        try {
            $raw = [System.Text.Encoding]::UTF8.GetString($bytes)
        } catch {
            $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
        }
    }

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
# PROCESS LOCAL ICS FILE LOGIC
# ============================================================

function Invoke-ProcessIcs {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) { return }

    $outlook = $null
    $namespace = $null
    $calendar = $null
    $createdCount = 0
    $updatedCount = 0
    $skippedCount = 0

    try {
        Write-Output ""
        Write-Output "[DETECTED] Processing local ICS file..."

        Start-Sleep -Seconds 2

        $outlook = Connect-Outlook
        $namespace = $outlook.GetNamespace("MAPI")
        $calendar = $namespace.GetDefaultFolder(9)

        $events = Read-IcsEvents -Path $FilePath

        Write-Output "Processing ICS events..."

        foreach ($icsEvent in $events) {
            $properties = $icsEvent.Properties
            $uid = $properties["UID"]
            if ([string]::IsNullOrWhiteSpace($uid)) { $skippedCount++; continue }

            $summary = Unescape-IcsText($properties["SUMMARY"])
            $description = Unescape-IcsText($properties["DESCRIPTION"])
            $location = Unescape-IcsText($properties["LOCATION"])
            $transp = $properties["TRANSP"]

            $targetBusyStatus = 2
            if ($summary -eq "Out of Office") {
                $targetBusyStatus = 3
            } elseif ($transp -eq "TRANSPARENT") {
                $targetBusyStatus = 0
            }

            try {
                $start = Parse-IcsDate($properties["DTSTART"])
                $end = Parse-IcsDate($properties["DTEND"])
            }
            catch {
                Write-Output "  [SKIP] Invalid date format for event: $summary"
                $skippedCount++
                continue
            }

            Write-Output "  -> Processing: '$summary' ($($start.ToString('yyyy-MM-dd')))..."

            $existing = Find-ExistingWorklogEvent -Calendar $calendar -UID $uid -Summary $summary -Start $start

            if ($null -eq $existing) {
                $appointment = $calendar.Items.Add(1)
                $appointment.Subject = $summary
                $appointment.Start = $start
                $appointment.End = $end
                $appointment.AllDayEvent = $true
                $appointment.Body = $description
                $appointment.Location = $location
                $appointment.BusyStatus = $targetBusyStatus

                $uidProperty = $appointment.UserProperties.Add("WorklogUID", 1, $false)
                $uidProperty.Value = $uid
                $appointment.Save()

                [Runtime.InteropServices.Marshal]::ReleaseComObject($uidProperty) | Out-Null
                [Runtime.InteropServices.Marshal]::ReleaseComObject($appointment) | Out-Null
                
                Write-Output " [CREATED]"
                $createdCount++
            }
            else {
                $normExistingBody = if ($null -eq $existing.Body) { "" } else { "$($existing.Body)".Replace("`r`n", "`n").Trim() }
                $normNewBody = if ($null -eq $description) { "" } else { "$description".Replace("`r`n", "`n").Trim() }

                $normExistingLoc = if ($null -eq $existing.Location) { "" } else { "$($existing.Location)".Trim() }
                $normNewLoc = if ($null -eq $location) { "" } else { "$location".Trim() }

                $existingStart = [datetime]$existing.Start
                $existingEnd = [datetime]$existing.End

                $changeReasons = @()
                if ($existing.Subject -ne $summary) { $changeReasons += "Subject" }
                if ($existingStart.Date -ne $start.Date) { $changeReasons += "Start Date" }
                if ($existingEnd.Date -ne $end.Date) { $changeReasons += "End Date" }
                if ($normExistingBody -ne $normNewBody) { $changeReasons += "Body" }
                if ($normExistingLoc -ne $normNewLoc) { $changeReasons += "Location" }
                if ($existing.BusyStatus -ne $targetBusyStatus) { $changeReasons += "BusyStatus" }

                $hasChanges = ($changeReasons.Count -gt 0)

                if ($hasChanges) {
                    $existing.Subject = $summary
                    $existing.Start = $start
                    $existing.End = $end
                    $existing.AllDayEvent = $true
                    $existing.Body = $description
                    $existing.Location = $location
                    $existing.BusyStatus = $targetBusyStatus

                    $uidProperty = $existing.UserProperties.Find("WorklogUID")
                    if ($null -eq $uidProperty) {
                        $uidProperty = $existing.UserProperties.Add("WorklogUID", 1, $false)
                    }
                    $uidProperty.Value = $uid
                    
                    $existing.Save()
                    if ($uidProperty) { [Runtime.InteropServices.Marshal]::ReleaseComObject($uidProperty) | Out-Null }
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($existing) | Out-Null

                    Write-Output " [UPDATED: $($changeReasons -join ', ')]"
                    $updatedCount++
                }
                else {
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($existing) | Out-Null
                    Write-Output " [SKIPPED - No Changes]"
                    $skippedCount++
                }
            }
        }

        Remove-Item $FilePath -Force -ErrorAction SilentlyContinue
        Write-Output "[OK] Local ICS file processed and cleaned up."
    }
    catch {
        Write-Output ""
        Write-Output "========================================"
        Write-Output "ERROR ENCOUNTERED:"
        Write-Output $_
        Write-Output "========================================"
    }
    finally {
        if ($calendar) { [Runtime.InteropServices.Marshal]::ReleaseComObject($calendar) | Out-Null }
        if ($namespace) { [Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null }
        if ($outlook) { [Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        if ($script:StartedOutlookByScript) {
            Stop-Process -Name OUTLOOK -Force -ErrorAction SilentlyContinue
        }

        Write-Output ""
        Write-Output "========================================"
        Write-Output "FINISHED. Created: $createdCount | Updated: $updatedCount | Skipped: $skippedCount"
        Write-Output "========================================"
    }
}

# ============================================================
# BACKGROUND POLLING LOOP (Replaces unreliable FileSystemWatcher)
# ============================================================

Write-Output ""
Write-Output "========================================"
Write-Output "BACKGROUND ICS POLLING ACTIVE"
Write-Output "Watching Folder: $LocalFolder"
Write-Output "========================================"
Write-Output "Keep this window minimized to run in background."

while ($true) {
    if (Test-Path $TargetIcsFile) {
        # Check if file is fully written/unlocked by OneDrive
        $ready = $false
        try {
            $stream = [System.IO.File]::Open($TargetIcsFile, 'Open', 'Read', 'None')
            $stream.Close()
            $stream.Dispose()
            $ready = $true
        } catch {
            # File is still syncing/locked by OneDrive
        }

        if ($ready) {
            Invoke-ProcessIcs -FilePath $TargetIcsFile
        }
    }
    Start-Sleep -Seconds 3
}