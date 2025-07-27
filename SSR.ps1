# Get the directory of the current script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Log file paths
$logFile = Join-Path $scriptDir "script.log"

# Define max log size (1MB)
$maxSizeBytes = 1MB

# Define log retention period (in days)
$logRetentionDays = 3

# Rotate log if it's too big
if (Test-Path $logFile) {
    $logSize = (Get-Item $logFile).Length
    if ($logSize -ge $maxSizeBytes) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $archivedLog = Join-Path $scriptDir "script_$timestamp.log"
        Rename-Item $logFile -NewName $archivedLog
    }
}

# Delete old logs (older than 3 days)
Get-ChildItem -Path $scriptDir -Filter "script_*.log" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$logRetentionDays) } |
    ForEach-Object {
        Remove-Item $_.FullName -Force
    }

# Define logger function
function Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$level] $message"
    Write-Output $logLine
    Add-Content -Path $logFile -Value $logLine
}

# Build dynamic source file path
$sourceFile = Join-Path $env:USERPROFILE "Desktop\NSB PHASE 2 FILES\NSB P2\NSB All files\SSR NSB P2 NEW.xlsx"

# Change working directory to script location
Set-Location $scriptDir
Log "Changed working directory to $scriptDir"

# Pull latest changes
Log "Running git pull..."
git pull *> $null
if ($LASTEXITCODE -eq 0) {
    Log "git pull completed successfully."
} else {
    Log "git pull failed." "ERROR"
}

# Define destination filename
$destFile = Join-Path $scriptDir "NSB-P2 SSR.xlsx"

# Copy and overwrite if exists
try {
    Copy-Item -Path $sourceFile -Destination $destFile -Force
    Log "Copied $sourceFile to $destFile"
} catch {
    Log "Failed to copy file: $_" "ERROR"
}

$module = "openpyxl"
$installed = python -c "import importlib.util; print(importlib.util.find_spec('$module') is not None)"

if ($installed -eq "False") {
    python -m pip install $module *> $null
}

# Git add
Log "Running git add..."
git add . *> $null

# Git commit
Log "Running git commit..."
git commit -m "Updated!" *> $null
if ($LASTEXITCODE -eq 0) {
    Log "git commit successful."
} else {
    Log "Nothing to commit or commit failed." "WARN"
}

# Git push
Log "Running git push..."
git push *> $null
if ($LASTEXITCODE -eq 0) {
    Log "git push successful."
} else {
    Log "git push failed." "ERROR"
}
