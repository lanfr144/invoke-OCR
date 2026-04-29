<#
.SYNOPSIS
    Installs the OCR Watcher as a background Scheduled Task and verifies all prerequisites.

.DESCRIPTION
    This script performs a complete installation of the Invoke-OCR background watcher service.
    It runs with automatic elevation to Administrator privileges.

    The installation process:
    1. Verifies PowerShell 7 (pwsh) is installed (required for parallel OCR processing)
    2. Checks that Tesseract, Ghostscript, and PDFtk are available on the system
    3. Creates the watch folder structure with language-specific subdirectories:
       - <WatchFolder>\en -> English (eng)
       - <WatchFolder>\de -> German (deu)
       - <WatchFolder>\fr -> French (fra)
       - <WatchFolder>\lb -> Luxembourgish (ltz)
    4. Registers a Windows Scheduled Task (InvokeOCR_Watcher) that:
       - Triggers at user logon
       - Runs hidden in the background
       - Persists through reboots
       - Works on battery power
    5. Starts the watcher service immediately

    The watcher monitors the folder and all subdirectories for new PDFs and images,
    automatically processing them with Invoke-OCR.ps1. You can customize processing
    per folder using .ocrconfig files (see Invoke-OCR.ps1 help for details).

.PARAMETER WatchFolder
    Root directory to monitor for incoming files. Default: C:\scans

.EXAMPLE
    .\Install-OCRWatcher.ps1

    Installs with default watch folder C:\scans.

.EXAMPLE
    .\Install-OCRWatcher.ps1 -WatchFolder "D:\incoming\scans"

    Installs with a custom watch directory.

.NOTES
    Requires: Administrator privileges (auto-elevates)
    Requires: PowerShell 7+ (pwsh), Tesseract, Ghostscript, PDFtk

    See also:
    - Get-OCRWatcherStatus.ps1 - Check if the watcher is running
    - Remove-OCRWatcher.ps1    - Uninstall the watcher
    - Start-OCRWatcher.ps1     - The watcher script itself

.LINK
    https://github.com/UB-Mannheim/tesseract/wiki

.LINK
    https://ghostscript.com/releases/gsdnld.html

.LINK
    https://www.pdflabs.com/tools/pdftk-server/
#>
param(
    [string]$WatchFolder = "C:\scans",
    [switch]$PassThru
)

$ErrorActionPreference = "Stop"

# Auto-elevate to Administrator (re-pass WatchFolder)
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Administrator privileges required. Requesting elevation..." -ForegroundColor Yellow
    $exe = (Get-Process -Id $PID).Path
    $passThruArg = if ($PassThru) { "-PassThru" } else { "" }
    Start-Process -FilePath $exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -WatchFolder `"$WatchFolder`" $passThruArg" -Verb RunAs
    return 0
}


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
    if ($PassThru) { [PSCustomObject]@{ ExitCode = 1; Message = "Prerequisites missing" } }
    return 1
}

# 3. Create Folders
Write-Highlight "`nSetting up directories..."
$baseFolder = $WatchFolder
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
$action = New-ScheduledTaskAction -Execute "pwsh.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$actionScript`" -WatchFolder `"$WatchFolder`""
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
Write-Host "You can now drop PDFs/images into $WatchFolder or its subfolders to automatically process them."

if ($PassThru) { [PSCustomObject]@{ ExitCode = 0; Message = "Success" } }
return 0
