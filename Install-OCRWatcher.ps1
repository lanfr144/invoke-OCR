<#
.SYNOPSIS
    Installs the OCR Watcher as a background Scheduled Task and verifies all prerequisites.
#>

$ErrorActionPreference = "Stop"

function Write-Highlight {
    param([string]$msg)
    Write-Host $msg -ForegroundColor Cyan
}

function Write-Success {
    param([string]$msg)
    Write-Host "[OK] $msg" -ForegroundColor Green
}

function Write-Fail {
    param([string]$msg)
    Write-Host "[X] $msg" -ForegroundColor Red
}

Write-Highlight "=== OCR Watcher Installer ==="
Write-Host "Checking prerequisites..."

$hasErrors = $false

# 1. Check PowerShell 7
if (Get-Command pwsh -ErrorAction SilentlyContinue) {
    Write-Success "PowerShell 7 (pwsh) found."
} else {
    Write-Fail "PowerShell 7 (pwsh) is NOT installed. Parallel OCR requires PowerShell 7."
    $hasErrors = $true
}

# 2. Check Programs
$programs = @(
    @{ Name="Tesseract"; Pattern="tesser*.exe" },
    @{ Name="Ghostscript"; Pattern="gswin*.exe" },
    @{ Name="PDFtk"; Pattern="pdftk*.exe" }
)

foreach ($prog in $programs) {
    $found = $false
    
    # Check PATH
    if (Get-Command $prog.Pattern -ErrorAction SilentlyContinue) {
        $found = $true
    } else {
        # Check Program Files
        $baseDirs = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Select-Object -Unique | Where-Object { $_ -ne $null }
        foreach ($baseDir in $baseDirs) {
            if (-not (Test-Path $baseDir)) { continue }
            $match = Get-ChildItem -Path $baseDir -Filter $prog.Pattern -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($match) {
                $found = $true
                break
            }
        }
    }

    if ($found) {
        Write-Success "$($prog.Name) found."
    } else {
        Write-Fail "$($prog.Name) is NOT installed."
        $hasErrors = $true
    }
}

if ($hasErrors) {
    Write-Host "`nInstallation aborted. Please install the missing prerequisites and try again." -ForegroundColor Red
    exit 1
}

# 3. Create Folders
Write-Highlight "`nSetting up directories..."
$baseFolder = "C:\scans"
$subFolders = @("en", "de", "fr", "lb")

if (-not (Test-Path $baseFolder)) {
    New-Item -ItemType Directory -Path $baseFolder | Out-Null
    Write-Success "Created $baseFolder"
} else {
    Write-Success "$baseFolder already exists"
}

foreach ($sub in $subFolders) {
    $subPath = Join-Path $baseFolder $sub
    if (-not (Test-Path $subPath)) {
        New-Item -ItemType Directory -Path $subPath | Out-Null
        Write-Success "Created $subPath"
    } else {
        Write-Success "$subPath already exists"
    }
}

# 4. Register Scheduled Task
Write-Highlight "`nRegistering Background Service..."
$taskName = "InvokeOCR_Watcher"

# Check if already exists
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Unregistering old task..."
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

$actionScript = Join-Path $PSScriptRoot "Start-OCRWatcher.ps1"

# Create action to run hidden powershell
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$actionScript`""
# Trigger on user logon
$trigger = New-ScheduledTaskTrigger -AtLogOn
# Run with highest privileges to avoid permission issues
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive
# Run regardless of AC power / laptop battery
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -DontStopOnIdleEnd -ExecutionTimeLimit 0

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings | Out-Null

Write-Success "Scheduled Task '$taskName' successfully registered!"
Write-Highlight "`nStarting the service for the first time..."
Start-ScheduledTask -TaskName $taskName

Write-Host "`nInstallation Complete! The background watcher is now running." -ForegroundColor Green
Write-Host "You can now drop PDFs/images into C:\scans or its subfolders to automatically process them."
