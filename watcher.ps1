# === watcher.ps1 (Fully Silent, Self-Cleaning, Debounced, Optimized) ===

# Get script directory
$scriptDir   = Split-Path -Parent $PSCommandPath
$scriptPath  = Join-Path $scriptDir "SSR.ps1"

# Source file to monitor
$sourceFile  = Join-Path $env:USERPROFILE "Desktop\NSB PHASE 2 FILES\NSB P2\NSB All files\SSR NSB P2 NEW.xlsx"
$folderPath  = Split-Path $sourceFile
$fileName    = Split-Path $sourceFile -Leaf

# Debounce config
$debounceSeconds      = 5
$global:LastTriggered = $null

# FileSystemWatcher setup
$watcher = [System.IO.FileSystemWatcher]::new()
$watcher.Path                = $folderPath
$watcher.Filter              = $fileName
$watcher.IncludeSubdirectories = $false
$watcher.NotifyFilter        = [System.IO.NotifyFilters]'LastWrite, FileName'

# Define silent action block with debounce
$action = {
    $now = Get-Date
    if ($global:LastTriggered -and ($now - $global:LastTriggered).TotalSeconds -lt $using:debounceSeconds) {
        return
    }
    $global:LastTriggered = $now

    try {
        & $using:scriptPath -sourceFile $using:sourceFile *> $null
    } catch {
        # Silently ignore errors
    }
}

# Register silent events
Register-ObjectEvent -InputObject $watcher -EventName Changed -SourceIdentifier "Watcher_Changed" -Action $action | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier "Watcher_Created" -Action $action | Out-Null
$watcher.EnableRaisingEvents = $true

# Cleanup on exit (Ctrl+C, window close, etc.)
Register-EngineEvent PowerShell.Exiting -Action {
    try {
        Unregister-Event -SourceIdentifier "Watcher_Changed" -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier "Watcher_Created" -ErrorAction SilentlyContinue
        $watcher.Dispose()
    } catch {}
} | Out-Null

# Infinite silent loop to keep watcher alive
while ($true) {
    Start-Sleep -Seconds 1
}