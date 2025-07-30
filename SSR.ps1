param (
    [Parameter(Position = 0)]
    [string]$sourceFile
)

# === Configuration ===
$scriptDir        = Split-Path -Parent $PSCommandPath
$logFile          = Join-Path $scriptDir "script.log"
$destFile         = Join-Path $scriptDir "NSB-P2 SSR.xlsx"
$push_ssr         = "FALSE"
$maxSizeBytes     = 1MB
$logRetentionDays = 3

# === Fallback for sourceFile ===
if (-not $sourceFile) {
    $sourceFile = Join-Path $env:USERPROFILE "Desktop\NSB PHASE 2 FILES\NSB P2\NSB All files\SSR NSB P2 NEW.xlsx"
}

# === Logging Function ===
function Log {
    param (
        [string]$message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$level] $message"
    Add-Content -Path $logFile -Value $logLine
}

# === Validate source file ===
function Validate-SourceFile {
    param ([string]$path)

    if (-not (Test-Path $path)) {
        Log "Source file does not exist: `"$path`"" "ERROR"
        exit 1
    }

    $ext = [IO.Path]::GetExtension($path)
    if ($ext -ne ".xlsx") {
        Log "Invalid file extension: '$ext'. Only .xlsx files are allowed." "ERROR"
        exit 1
    }
}

# === Log rotation ===
if (Test-Path $logFile) {
    $logSize = (Get-Item $logFile).Length
    if ($logSize -ge $maxSizeBytes) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $archivedLog = Join-Path $scriptDir "script_$timestamp.log"
        Rename-Item -Path $logFile -NewName $archivedLog
    }
}

# === Log cleanup ===
Get-ChildItem -Path $scriptDir -Filter "script_*.log" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$logRetentionDays) } |
    Remove-Item -Force

# === Begin execution ===
Validate-SourceFile -path $sourceFile

Set-Location $scriptDir
Log "Changed working directory to $scriptDir"

# Git pull
Log "Running git pull..."
$gitPullResult = git pull 2>&1
$gitPullResult | ForEach-Object { Log $_ }
if ($LASTEXITCODE -eq 0) {
    Log "git pull completed successfully."
} else {
    Log "git pull failed." "ERROR"
}

# Copy file
try {
    Copy-Item -Path $sourceFile -Destination $destFile -Force
    Log "Copied `"$sourceFile`" to `"$destFile`""
    $push_ssr = "TRUE"
} catch {
    Log "Failed to copy file: $_" "ERROR"
}

# Persist PUSH_SSR env var for the job
[Environment]::SetEnvironmentVariable("PUSH_SSR", $push_ssr, "User")

$jobName = "SSR_AutoPush"

# Kill old job if exists
if (Get-Job -Name $jobName -ErrorAction SilentlyContinue) {
    Remove-Job -Name $jobName -Force
}

# Start Background Git Job
Start-Job -Name $jobName -ScriptBlock {
    function HasInternet {
        try {
            $null = Invoke-RestMethod -Uri "https://www.google.com" -TimeoutSec 3
            return $true
        } catch {
            return $false
        }
    }

    function GitHasChanges {
        $status = git status --porcelain
        return -not [string]::IsNullOrWhiteSpace($status)
    }

    while ($true) {
        $envVal = [Environment]::GetEnvironmentVariable("PUSH_SSR", "User")
        if ($envVal -eq "TRUE") {
            if (GitHasChanges) {
                if (HasInternet) {
                    try {
                        git add . 2>&1 | Out-Null
                        git commit -m "Auto-commit by background service" --no-verify 2>&1 | Out-Null
                        git push 2>&1 | Out-Null
                        [Environment]::SetEnvironmentVariable("PUSH_SSR", "FALSE", "User")
                    } catch {
                        Start-Sleep -Seconds 5
                    }
                }
            } else {
                [Environment]::SetEnvironmentVariable("PUSH_SSR", "FALSE", "User")
            }
        } else {
            break
        }
        Start-Sleep -Seconds 30
    }
}