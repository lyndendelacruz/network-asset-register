# ============================
# sync.ps1 — LIVE GITHUB SYNC
# ============================

# ----------------------------
# CONFIGURATION
# ----------------------------

# Path to your SharePoint‑synced Excel file
$excelPath = "C:\Users\Lynden.DeLaCruz\Cloud Direct\Network Site Info Update - Asset Register\JCL-VM-Asset-Register-Live-Ext.xlsx"

# Path to your GitHub repo (UPDATED)
$repoPath = "C:\Users\Lynden.DeLaCruz\OneDrive - Cloud Direct\Desktop\dev\network-asset-register"

# Path to your Node converter script (UPDATED)
$convertScript = "$repoPath\scripts\convert.js"

# Debounce delay (ms)
$debounce = 1500


# ----------------------------
# WATCHER SETUP
# ----------------------------

Write-Host "Starting watcher..."
Write-Host "Watching: $excelPath"
Write-Host "Repo: $repoPath"
Write-Host "-------------------------`n"

$fsw = New-Object System.IO.FileSystemWatcher
$fsw.Path = (Split-Path $excelPath)
$fsw.Filter = (Split-Path $excelPath -Leaf)

# Excel + OneDrive require multiple notify filters
$fsw.NotifyFilter = [System.IO.NotifyFilters]'LastWrite, FileName, Size'


# ----------------------------
# DEBOUNCE TIMER
# ----------------------------

$timer = New-Object System.Timers.Timer
$timer.Interval = $debounce
$timer.AutoReset = $false


# ----------------------------
# ACTION ON CHANGE
# ----------------------------

$action = {
    Write-Host "`nChange detected. Running conversion..."

    # Run Node converter
    node $convertScript | Write-Host

    # Commit + push
    Set-Location $repoPath
    git add .
    git commit -m "Auto-update CSV $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" --allow-empty
    git push

    Write-Host "Sync complete."
}

$timerEvent = Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action $action


# ----------------------------
# FILE CHANGE EVENT
# ----------------------------

Register-ObjectEvent -InputObject $fsw -EventName Changed -Action {
    $timer.Stop()
    $timer.Start()
}


# ----------------------------
# KEEP SCRIPT ALIVE
# ----------------------------

while ($true) { Start-Sleep -Seconds 1 }