$ErrorActionPreference = "Stop"

$outlook = $null
$namespace = $null
$calendar = $null
$items = $null
$event = $null

try {

    Write-Host "Connecting to Outlook..."

    $outlook = New-Object -ComObject Outlook.Application

    $namespace = $outlook.GetNamespace("MAPI")

    $calendar = $namespace.GetDefaultFolder(9)

    Write-Host "Calendar: $($calendar.Name)"
    Write-Host ""
    Write-Host "Creating test event..."

    $event = $calendar.Items.Add(1)

    $event.Subject = "WORKLOG AUTOMATION TEST"
    $event.Start = (Get-Date).AddMinutes(10)
    $event.End = (Get-Date).AddMinutes(40)
    $event.BusyStatus = 0
    $event.Body = "Temporary test event created by PowerShell."

    $event.Save()

    Write-Host "✓ Test event created."
    Write-Host ""
    Write-Host "Now deleting it..."

    $event.Delete()

    Write-Host "✓ Test event deleted."
    Write-Host ""
    Write-Host "SUCCESS: PowerShell can create and modify Outlook calendar events."

}
catch {

    Write-Host ""
    Write-Host "FAILED:"
    Write-Host $_.Exception.Message
}
finally {

    if ($event) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($event) | Out-Null
    }

    if ($items) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null
    }

    if ($calendar) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($calendar) | Out-Null
    }

    if ($namespace) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
    }

    if ($outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}