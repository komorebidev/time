# ============================================================
# START OUTLOOK MINIMIZED (VIA COM ACTIVE WINDOW STATE)
# ============================================================

function Start-Outlook-Minimized {
    Write-Host ""
    Write-Host "Starting Outlook minimized..."

    # Launch normally via shell first to spawn process
    Start-Process "outlook.exe"
    $script:StartedOutlookByScript = $true

    # Allow application to boot up, then hook via COM and force minimize window state
    $maxAttempts = 15
    for ($i = 1; $i -le $maxAttempts; $i++) {
        Start-Sleep -Seconds 1
        try {
            $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            if ($app -and $app.Explorers.Count -gt 0) {
                $activeWin = $app.ActiveWindow()
                if ($activeWin) {
                    $activeWin.WindowState = 1 # olMinimized = 1
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
                    break
                }
            }
            if ($app) { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null }
        }
        catch {
            # Retry if COM isn't fully responsive yet
        }
    }
    Start-Sleep -Seconds 2
}