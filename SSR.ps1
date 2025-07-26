# Build dynamic source file path
$sourceFile = Join-Path $env:USERPROFILE "Desktop\NSB PHASE 2 FILES\NSB P2\NSB All files\SSR NSB P2 NEW.xlsx"

# Get the directory of the current script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Change working directory to script location
Set-Location $scriptDir

git pull *> $null

# Define destination filename
$destFile = Join-Path $scriptDir "SSR.xlsx"

# Copy and overwrite if exists
Copy-Item -Path $sourceFile -Destination $destFile -Force

# Perform Git operations silently
git add . *> $null
git commit -m "Updated!" *> $null
git push *> $null
