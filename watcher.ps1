# watcher.ps1 (fully silent, optimized)

$scriptDir = Split-Path -Parent $PSCommandPath

# Define file to monitor
$sourceFile = Join-Path $env:USERPROFILE "Desktop\NSB PHASE 2 FILES\NSB P2\NSB All files\SSR NSB P2 NEW.xlsx"
$folderPath = Split-Path $sourceFile
$fileName = Split-Path $sourceFile -Leaf

# Setup watcher
$watcher = [System.IO.FileSystemWatcher]::new()
$watcher.Path = $folderPath
$watcher.Filter = $fileName
$watcher.IncludeSubdirectories = $false
$watcher.NotifyFilter = [System.IO.NotifyFilters]'LastWrite, FileName'

# Debounce mechanism
$global:LastTriggered = $null
$debounceSeconds = 5

# Handler
$action = {
    $now = Get-Date
    if ($global:LastTriggered -and ($now - $global:LastTriggered).TotalSeconds -lt $using:debounceSeconds) {
        return
    }

    $global:LastTriggered = $now
    $scriptPath = Join-Path $using:scriptDir "SSR.ps1"

    try {
        & $scriptPath -sourceFile $using:sourceFile *>$null
    } finally {
        # No output, no catch, just silent cleanup if needed
    }
}

# Register silent events
$null = Register-ObjectEvent -InputObject $watcher -EventName Changed -SourceIdentifier FileChanged -Action $action
$null = Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier FileCreated -Action $action

$watcher.EnableRaisingEvents = $true

# Prevent terminal output forever
while ($true) {
    Start-Sleep -Seconds 1
}
