<#
.SYNOPSIS
    Removes the InvokeOCR_Watcher Scheduled Task from the system.

.DESCRIPTION
    Completely uninstalls the Invoke-OCR background watcher service by:
    1. Stopping the running task (if active)
    2. Waiting for graceful shutdown
    3. Unregistering the Scheduled Task from Windows Task Scheduler

    This does NOT remove the C:\scans folder structure or any processed files.
    Requires Administrator privileges (auto-elevates if needed).

.EXAMPLE
    .\Remove-OCRWatcher.ps1

    Stops and unregisters the watcher. Will auto-request Administrator elevation if needed.

.NOTES
    Requires: Administrator privileges (auto-elevates)

    See also:
    - Install-OCRWatcher.ps1   - Re-install the watcher
    - Get-OCRWatcherStatus.ps1 - Check current status
#>

$taskName = "InvokeOCR_Watcher"

# Auto-elevate to Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Administrator privileges required. Requesting elevation..." -ForegroundColor Yellow
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

if (-not $task) {
    Write-Host "The OCR Watcher service is not installed. Nothing to remove." -ForegroundColor Yellow
} else {
    Write-Host "Stopping service..."
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    Write-Host "Unregistering Scheduled Task..."
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    
    Write-Host "Service completely removed!" -ForegroundColor Green
}
