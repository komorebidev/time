$ErrorActionPreference = "Stop"

$outlook = $null
$namespace = $null
$inbox = $null
$items = $null

try {

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items

    Write-Host ""
    Write-Host "========================================"
    Write-Host "OUTLOOK INBOX TEST"
    Write-Host "========================================"
    Write-Host ""

    Write-Host "Inbox:"
    Write-Host $inbox.FolderPath
    Write-Host ""

    Write-Host "Total items:"
    Write-Host $items.Count
    Write-Host ""

    Write-Host "Latest messages:"
    Write-Host "----------------------------------------"

    $items.Sort("[ReceivedTime]", $true)

    $count = 0

    foreach ($item in $items) {

        if ($item.Class -ne 43) {
            continue
        }

        $count++

        Write-Host ""
        Write-Host "Subject: [$($item.Subject)]"
        Write-Host "From:    [$($item.SenderEmailAddress)]"
        Write-Host "Date:    [$($item.ReceivedTime)]"
        Write-Host "Files:   $($item.Attachments.Count)"

        if ($count -ge 20) {
            break
        }
    }

    Write-Host ""
    Write-Host "========================================"

}
catch {

    Write-Host ""
    Write-Host "FAILED:"
    Write-Host $_.Exception.Message
}
finally {

    if ($items) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null
    }

    if ($inbox) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($inbox) | Out-Null
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