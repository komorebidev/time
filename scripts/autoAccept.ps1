# ============================================================
# Worklog ICS -> Outlook Calendar (OneDrive Watcher / Pure COM)
# ============================================================

$ErrorActionPreference = "Stop"

# Force UTF-8 encoding for console display to avoid mojibake
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

# ============================================================
# CONFIGURATION
# ============================================================

$LocalFolder = "$env:USERPROFILE\OneDrive - エイラシステム株式会社\ics"
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
            Write-Host "[OK] Connected to active Outlook instance"
            break
        }
        catch {
            try {
                $outlookType = [Type]::GetTypeFromProgID("Outlook.Application")
                $outlook = [Activator]::CreateInstance($outlookType)
                $script:StartedOutlookByScript = $true
                Write-Host "[OK] Created background Outlook COM instance"
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
# READ ICS EVENTS (Handles UTF-8 and Shift-JIS automatically)
# ============================================================

function Read-IcsEvents {
    param([string]$Path)
    
    # Read bytes first to check for BOM or fallback safely
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $encoding = [System.Text.Encoding]::UTF8
    
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $encoding = [System.Text.Encoding]::UTF8
    } else {
        # Test if valid UTF-8 string, otherwise fallback to system default (Shift-JIS for Japanese Windows)
        try {
            $utf8String = [System.Text.Encoding]::UTF8.GetString($bytes)
            # Simple heuristic or default fallback
            $encoding = [System.Text.Encoding]::UTF8
        } catch {
            $encoding = [System.Text.Encoding]::Default
        }
    }

    $raw = [System.Text.Encoding]::UTF8.GetString($bytes)
    # If standard text read fails or has invalid chars, fallback to Default
    try {
        $raw = $encoding.GetString($bytes)
    } catch {
        $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
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
        $partsCount = $parts.Count
        if ($partsCount -ne 2) { continue }

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
        Write-Host ""
        Write-Host "[DETECTED] Processing local ICS file..." -ForegroundColor Cyan

        # Brief pause to ensure OneDrive has fully released the sync write lock
        Start-Sleep -Seconds 3

        $outlook = Connect-Outlook
        $namespace = $outlook.GetNamespace("MAPI")
        $calendar = $namespace.GetDefaultFolder(9)

        $events = Read-IcsEvents -Path $FilePath

        Write-Host ""
        Write-Host "Processing ICS events..."

        foreach ($icsEvent in $events) {
            $properties = $icsEvent.Properties
            $uid = $properties["UID"]
            if ([string]::IsNullOrWhiteSpace($uid)) { $skippedCount++; continue }

            $summary = Unescape-IcsText($properties["SUMMARY"])
            $description = Unescape-IcsText($properties["DESCRIPTION"])
            $location = Unescape-IcsText($properties["LOCATION"])
            $transp = $properties["TRANSP"]

            # Determine BusyStatus: Out of Office = 3, Transparent (Free) = 0, Regular Busy = 2
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
                Write-Host "  [SKIP] Invalid date format for event: $summary" -ForegroundColor Yellow
                $skippedCount++
                continue
            }

            Write-Host "  -> Processing: '$summary' ($($start.ToString('yyyy-MM-dd')))..." -NoNewline

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
                
                Write-Host " [CREATED]" -ForegroundColor Green
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

                    Write-Host " [UPDATED: $($changeReasons -join ', ')]" -ForegroundColor Cyan
                    $updatedCount++
                }
                else {
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($existing) | Out-Null
                    Write-Host " [SKIPPED - No Changes]" -ForegroundColor DarkGray
                    $skippedCount++
                }
            }
        }

        # Delete local ICS file after successful processing
        Remove-Item $FilePath -Force -ErrorAction SilentlyContinue
        Write-Host "[OK] Local ICS file processed and cleaned up." -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Host "========================================"
        Write-Host "ERROR ENCOUNTERED:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        Write-Host "========================================"
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

        Write-Host ""
        Write-Host "========================================"
        Write-Host "FINISHED. Created: $createdCount | Updated: $updatedCount | Skipped: $skippedCount"
        Write-Host "========================================"
    }
}

# ============================================================
# FILE SYSTEM WATCHER SETUP (Background Listener)
# ============================================================

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $LocalFolder
$watcher.Filter = $Filter
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

Write-Host ""
Write-Host "========================================"
Write-Host "BACKGROUND ICS WATCHER ACTIVE"
Write-Host "Watching: $LocalFolder"
Write-Host "========================================"
Write-Host "Keep this window minimized to run in background."

$action = {
    $path = $Event.SourceEventArgs.FullPath
    
    # Verify file is unlocked and fully synced
    $success = $false
    for ($i = 1; $i -le 5; $i++) {
        try {
            if (Test-Path $path) {
                $stream = [System.IO.File]::Open($path, 'Open', 'Read', 'None')
                $stream.Close()
                $stream.Dispose()
                $success = $true
                break
            }
        } catch {
            Start-Sleep -Seconds 2
        }
    }

    if ($success) {
        Invoke-ProcessIcs -FilePath $path
    }
}

Register-ObjectEvent $watcher "Created" -Action $action | Out-Null
Register-ObjectEvent $watcher "Changed" -Action $action | Out-Null

# Catch any file already sitting in the folder upon startup
if (Test-Path $TargetIcsFile) {
    Write-Host "[INFO] Existing file detected on startup. Processing..." -ForegroundColor Yellow
    Invoke-ProcessIcs -FilePath $TargetIcsFile
}

try {
    while ($true) { Start-Sleep -Seconds 1 }
}
finally {
    Unregister-Event -SourceIdentifier *
    $watcher.Dispose()
}