# === watcher.ps1 (Silent, Debounced, Background-Compatible, Logging) ===

# Get script directory and paths
$scriptDir = Split-Path -Parent $PSCommandPath
$scriptPath = Join-Path $scriptDir "SSR.ps1"
$logPath = Join-Path $scriptDir "script.log"

# Source file to monitor
$sourceFile = Join-Path $env:USERPROFILE "Desktop\NSB PHASE 2 FILES\NSB P2\NSB All files\SSR NSB P2 NEW.xlsx"
$folderPath = Split-Path $sourceFile -Parent
$fileName = Split-Path $sourceFile -Leaf

# Debounce settings
$debounceSeconds = 5
$global:LastTriggered = $null

# Create watcher
$watcher = [System.IO.FileSystemWatcher]::new()
$watcher.Path = $folderPath
$watcher.Filter = $fileName
$watcher.IncludeSubdirectories = $false
$watcher.NotifyFilter = [System.IO.NotifyFilters]'LastWrite, FileName'

# Logging helper
function Log {
    param ([string]$message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp $message" | Out-File -FilePath $logPath -Append -Encoding utf8
}

# Debounced action block
$action = {
    $now = Get-Date
    if ($global:LastTriggered -and ($now - $global:LastTriggered).TotalSeconds -lt $using:debounceSeconds) {
        return
    }
    $global:LastTriggered = $now

    try {
        Log "Change detected. Running SSR.ps1..."
        & $using:scriptPath -sourceFile $using:sourceFile *> $null
        Log "SSR.ps1 completed successfully."
    }
    catch {
        Log "Error while running SSR.ps1: $_"
    }
}

# Register silent events
Register-ObjectEvent -InputObject $watcher -EventName Changed -SourceIdentifier "Watcher_Changed" -Action $action | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier "Watcher_Created" -Action $action | Out-Null
$watcher.EnableRaisingEvents = $true

# Cleanup on exit
Register-EngineEvent PowerShell.Exiting -Action {
    try {
        Unregister-Event -SourceIdentifier "Watcher_Changed" -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier "Watcher_Created" -ErrorAction SilentlyContinue
        $watcher.Dispose()
        Log "Watcher exited."
    }
    catch {}
} | Out-Null

Log "Watcher started."

# Keep running silently
while ($true) {
    Start-Sleep -Seconds 1
}
