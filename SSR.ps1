param (
    [Parameter(Position = 0)]
    [string]$sourceFile,

    [Parameter(Position = 1)]
    [Alias("msg")]
    [string]$commitMessage = "Update! Auto-commit by background service"
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

# === Git pull ===
Log "Running git pull..."
$gitPullResult = git pull 2>&1
$gitPullResult | ForEach-Object { Log $_ }
if ($LASTEXITCODE -eq 0) {
    Log "git pull completed successfully."
} else {
    Log "git pull failed." "ERROR"
}

# === Copy file ===
try {
    Copy-Item -Path $sourceFile -Destination $destFile -Force
    Log "Copied `"$sourceFile`" to `"$destFile`""
} catch {
    Log "Failed to copy file: $_" "ERROR"
    exit 1
}

# === Python/Package Bootstrapping ===

function Is-CommandAvailable($cmd) {
    return (Get-Command $cmd -ErrorAction SilentlyContinue) -ne $null
}

function Install-Choco {
    Log "Installing Chocolatey..."
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-Expression ((New-Object Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
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

# Refresh PATH (important if Python was just installed)
$env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
            [Environment]::GetEnvironmentVariable("Path", "User")

# Detect python or python3
if (Is-CommandAvailable "python") {
    $pythonCmd = "python"
} elseif (Is-CommandAvailable "python3") {
    $pythonCmd = "python3"
} else {
    Install-Python
    $pythonCmd = "python"
}

# Check pip
if (-not (& $pythonCmd -m pip --version 2>$null)) {
    Log "pip not found. Installing via get-pip.py..."
    Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -OutFile get-pip.py
    & $pythonCmd get-pip.py
    Remove-Item get-pip.py
}

# Check xlwings
if (-not (& $pythonCmd -m pip show xlwings 2>$null)) {
    Log "Installing xlwings..."
    & $pythonCmd -m pip install xlwings
} else {
    Log "xlwings is already installed."
}

# === Run resolver ===
if (-not (Test-Path $resolverScript)) {
    Log "resolver.py not found at $resolverScript" "ERROR"
    exit 1
}

Log "Running resolver.py..."
& $pythonCmd $resolverScript

# === Git push ===
git add . 2>&1 | Out-Null
git commit -m "$commitMessage" --no-verify 2>&1 | Out-Null
git push 2>&1 | Out-Null
Log "Git auto-push completed."
