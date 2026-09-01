# ============================================================
# Worklog ICS -> Outlook Calendar (FileSystemWatcher / Pure COM)
# ============================================================

$ErrorActionPreference = "Stop"

# Force UTF-8 encoding for console output and standard streams to prevent mojibake
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ============================================================
# CONFIGURATION (Dynamically locates OneDrive to avoid dash/encoding bugs)
# ============================================================

$BaseOneDrive = Get-ChildItem "$env:USERPROFILE\OneDrive*" | Where-Object { $_.Name -like "*エイラシステム*" } | Select-Object -ExpandProperty FullName

if (-not $BaseOneDrive) {
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
# SAFE COM RELEASE HELPER
# ============================================================

function Release-ComObject {
    param($ComObject)
    if ($null -ne $ComObject) {
        try {
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
            }
        }
        catch {}
    }
}

# ============================================================
# CONNECT TO OUTLOOK (BULLETPROOF COM INSTANTIATION)
# ============================================================

function Connect-Outlook {
    Write-Output ""
    Write-Output "Connecting to Outlook via COM..."

    $outlook = $null
    $maxAttempts = 10

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
            
            if ($outlookProcess) {
                try {
                    $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
                } catch {
                    $outlook = New-Object -ComObject Outlook.Application
                }
            } else {
                $outlook = New-Object -ComObject Outlook.Application
                $script:StartedOutlookByScript = $true
            }

            if ($null -ne $outlook) {
                $ns = $outlook.GetNamespace("MAPI")
                if ($null -ne $ns) {
                    Release-ComObject $ns
                    Write-Output "[OK] Connected to Outlook successfully."
                    break
                }
            }
        }
        catch {
            Write-Output "Attempt $attempt of $maxAttempts failed to bind Outlook COM. Retrying..."
            if ($attempt -eq $maxAttempts) {
                throw "Failed to initialize Outlook COM interface: $_"
            }
            Start-Sleep -Seconds 3
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
                if ($null -eq $item) { continue }
                if ($item.Class -ne 26) { 
                    Release-ComObject $item
                    continue 
                }

                $property = $null
                try {
                    $property = $item.UserProperties.Find("WorklogUID")
                    if ($null -ne $property -and $property.Value -eq $UID) { return $item }
                }
                finally {
                    Release-ComObject $property
                }

                if ($item.Subject -eq $Summary) {
                    $itemStart = [datetime]$item.Start
                    if ($itemStart.Date -eq $Start.Date) { return $item }
                }
                
                Release-ComObject $item
            }
            catch {
                Release-ComObject $item
            }
        }
    }
    finally {
        Release-ComObject $items
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

        Start-Sleep -Seconds 3

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

                Release-ComObject $uidProperty
                Release-ComObject $appointment
                
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
                    Release-ComObject $uidProperty
                    Release-ComObject $existing

                    Write-Output " [UPDATED: $($changeReasons -join ', ')]"
                    $updatedCount++
                }
                else {
                    Release-ComObject $existing
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
        Release-ComObject $calendar
        Release-ComObject $namespace
        Release-ComObject $outlook

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
# FILE SYSTEM WATCHER SETUP
# ============================================================

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $LocalFolder
$watcher.Filter = $Filter
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

Write-Output ""
Write-Output "========================================"
Write-Output "BACKGROUND ICS WATCHER ACTIVE"
Write-Output "Watching Folder: $LocalFolder"
Write-Output "========================================"
Write-Output "Keep this window minimized to run in background."

$action = {
    $path = $Event.SourceEventArgs.FullPath
    
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

if (Test-Path $TargetIcsFile) {
    Write-Output "[INFO] Existing file detected on startup. Processing..."
    Invoke-ProcessIcs -FilePath $TargetIcsFile
}

try {
    while ($true) { Start-Sleep -Seconds 1 }
}
finally {
    Unregister-Event -SourceIdentifier *
    $watcher.Dispose()
}