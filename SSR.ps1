param (
    [Parameter(Position = 0)]
    [string]$sourceFile
)

# === Configuration ===
$scriptDir        = Split-Path -Parent $PSCommandPath
$resolverScript   = Join-Path $scriptDir "resolver.py"
$logFile          = Join-Path $scriptDir "script.log"
$destFile         = Join-Path $scriptDir "NSB-P2 SSR.xlsx"
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
} catch {
    Log "Failed to copy file: $_" "ERROR"
    exit 1
}

function Is-CommandAvailable($cmd) {
    return (Get-Command $cmd -ErrorAction SilentlyContinue) -ne $null
}

function Install-Choco {
    Log "Installing Chocolatey..."
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
}

function Install-Python {
    if (Is-CommandAvailable "winget") {
        Log "Installing Python via winget..."
        winget install --id Python.Python.3 --source winget -e
    } elseif (Is-CommandAvailable "choco") {
        Log "Installing Python via Chocolatey..."
        choco install -y python
    } else {
        Log "No package manager found. Installing Chocolatey..."
        Install-Choco
        Log "Installing Python via Chocolatey..."
        choco install -y python
    }
}

# Step 1: Ensure Python is installed
if (-not (Is-CommandAvailable "python")) {
    Install-Python
} else {
    Log "Python is already installed."
}

# Step 2: Ensure pip is available
try {
    & python -m pip --version > $null
} catch {
    Log "pip not found. Installing via get-pip.py..."
    Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -OutFile get-pip.py
    python get-pip.py
    Remove-Item get-pip.py
}

# Step 3: Check and install xlwings
if (-not (python -m pip show xlwings 2>$null)) {
    Log "Installing xlwings..."
    python -m pip install xlwings
} else {
    Log "xlwings is already installed."
}

python $resolverScript

git add . 2>&1 | Out-Null
git commit -m "Auto-commit by background service" --no-verify 2>&1 | Out-Null
git push 2>&1 | Out-Null

# Persist PUSH_SSR env var for the job
# [Environment]::SetEnvironmentVariable("PUSH_SSR", "TRUE", "User")

# $jobName = "SSR_AutoPush"

# # Kill old job if exists
# if (Get-Job -Name $jobName -ErrorAction SilentlyContinue) {
#     Remove-Job -Name $jobName -Force
# }

# # Start Background Git Job
# Start-Job -Name $jobName -ScriptBlock {
#     function HasInternet {
#         try {
#             $null = Invoke-RestMethod -Uri "https://www.google.com" -TimeoutSec 3
#             return $true
#         } catch {
#             return $false
#         }
#     }

#     function GitHasChanges {
#         $status = git status --porcelain
#         return -not [string]::IsNullOrWhiteSpace($status)
#     }

#     while ($true) {
#         $envVal = [Environment]::GetEnvironmentVariable("PUSH_SSR", "User")
#         if ($envVal -eq "TRUE") {
#             if (GitHasChanges) {
#                 if (HasInternet) {
#                     try {
#                         git add . 2>&1 | Out-Null
#                         git commit -m "Auto-commit by background service" --no-verify 2>&1 | Out-Null
#                         git push 2>&1 | Out-Null
#                         [Environment]::SetEnvironmentVariable("PUSH_SSR", "FALSE", "User")
#                     } catch {
#                         Start-Sleep -Seconds 5
#                     }
#                 }
#             } else {
#                 [Environment]::SetEnvironmentVariable("PUSH_SSR", "FALSE", "User")
#             }
#         } else {
#             break
#         }
#         Start-Sleep -Seconds 30
#     }
# }