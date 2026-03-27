# ================================
#  Network Asset Register Watcher
#  5-second debounce, OneDrive-safe
#  Commit message: Auto-update: synced Excel changes
# ================================

$watchPath = "C:\Users\Lynden.DeLaCruz\OneDrive - Cloud Direct\Desktop\dev\network-asset-register"
$debounceSeconds = 5
$timer = New-Object Timers.Timer
$timer.Interval = $debounceSeconds * 1000
$timer.AutoReset = $false

Write-Host "Watcher started. Monitoring:" -ForegroundColor Cyan
Write-Host $watchPath -ForegroundColor Yellow

# --- Helper: Wait until file is fully written and OneDrive is done syncing ---
function Wait-ForStableFile {
    param([string]$filePath)

    while ($true) {
        try {
            $stream = [System.IO.File]::Open($filePath, 'Open', 'Read', 'None')
            $stream.Close()
            break
        } catch {
            Start-Sleep -Milliseconds 300
        }
    }

    # Wait for OneDrive to finish syncing (no .tmp, .partial, .lock)
    Start-Sleep -Milliseconds 500
}

# --- Debounced action ---
$action = {
    Write-Host "Quiet period reached. Processing changes..." -ForegroundColor Cyan

    # Ensure repo is clean
    git -C $watchPath add .
    git -C $watchPath commit -m "Auto-update: synced Excel changes"
    git -C $watchPath push

    Write-Host "GitHub updated successfully." -ForegroundColor Green
}

# --- FileSystemWatcher setup ---
$fsw = New-Object IO.FileSystemWatcher $watchPath, "*.xlsx"
$fsw.IncludeSubdirectories = $false
$fsw.EnableRaisingEvents = $true

# --- Event handler ---
$handler = Register-ObjectEvent $fsw Changed -Action {
    $file = $Event.SourceEventArgs.FullPath

    if ($file -match "~\$") { return }   # Ignore temp Excel files

    Write-Host "Change detected: $file" -ForegroundColor Magenta

    Wait-ForStableFile $file

    # Reset debounce timer
    $timer.Stop()
    $timer.Start()
}

# --- Timer event ---
Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action $action

# Keep script alive
while ($true) { Start-Sleep 1 }