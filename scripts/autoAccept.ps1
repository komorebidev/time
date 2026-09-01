$ErrorActionPreference = "Stop"

# ============================================================
# CONFIGURATION
# ============================================================
$LocalFolder = "$env:USERPROFILE\OneDrive - エイラシステム株式会社\ics"
$Filter = "time.ics"

if (-not (Test-Path $LocalFolder)) {
    New-Item -ItemType Directory -Path $LocalFolder -Force | Out-Null
}

$script:StartedOutlookByScript = $false

# ============================================================
# OUTLOOK COM & ICS FUNCTIONS
# ============================================================

function Connect-Outlook {
    $outlook = $null
    $maxAttempts = 10

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") -as [Microsoft.Office.Interop.Outlook.Application]
            if ($null -ne $outlook) { break }
            throw "No active instance found."
        }
        catch {
            try {
                $outlookType = [Type]::GetTypeFromProgID("Outlook.Application")
                $outlook = [Activator]::CreateInstance($outlookType)
                $script:StartedOutlookByScript = $true
                break
            }
            catch {
                if ($attempt -eq $maxAttempts) {
                    throw "Failed to initialize Outlook COM interface: $_"
                }
                Start-Sleep -Seconds 3
            }
        }
    }
    return $outlook
}

function Unescape-IcsText {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    $Text = $Text -replace '\\n', "`r`n" -replace '\\N', "`r`n" -replace '\\,', "," -replace '\\;', ";" -replace '\\\\', "\"
    return $Text
}

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

function Invoke-ProcessIcs {
    param([string]$FilePath)

    $outlook = $null
    $namespace = $null
    $calendar = $null
    $createdCount = 0
    $updatedCount = 0
    $skippedCount = 0

    try {
        Write-Host "[DETECTED] Processing local ICS file..." -ForegroundColor Cyan

        $outlook = Connect-Outlook
        $namespace = $outlook.GetNamespace("MAPI")
        $calendar = $namespace.GetDefaultFolder(9)

        $events = Read-IcsEvents -Path $FilePath

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
                $appointment.BusyStatus = $targetBusyStatus

                $uidProperty = $appointment.UserProperties.Add("WorklogUID", 1, $false)
                $uidProperty.Value = $uid
                $appointment.Save()

                [Runtime.InteropServices.Marshal]::ReleaseComObject($uidProperty) | Out-Null
                [Runtime.InteropServices.Marshal]::ReleaseComObject($appointment) | Out-Null
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

                if ($changeReasons.Count -gt 0) {
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
                    $updatedCount++
                }
                else {
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($existing) | Out-Null
                    $skippedCount++
                }
            }
        }

        # Delete file after successful import
        Remove-Item $FilePath -Force -ErrorAction SilentlyContinue
        Write-Host "[SUCCESS] Imported. Created: $createdCount | Updated: $updatedCount | Skipped: $skippedCount" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] $_" -ForegroundColor Red
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

Write-Host "========================================"
Write-Host "BACKGROUND ICS WATCHER ACTIVE"
Write-Host "Watching: $LocalFolder"
Write-Host "========================================"

$action = {
    $path = $Event.SourceEventArgs.FullPath
    Start-Sleep -Seconds 3 # Wait for OneDrive sync lock to clear

    $success = $false
    for ($i = 1; $i -le 5; $i++) {
        try {
            $stream = [System.IO.File]::Open($path, 'Open', 'Read', 'None')
            $stream.Close()
            $stream.Dispose()
            $success = $true
            break
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

try {
    while ($true) { Start-Sleep -Seconds 1 }
}
finally {
    Unregister-Event -SourceIdentifier *
    $watcher.Dispose()
}