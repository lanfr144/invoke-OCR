<#
.SYNOPSIS
    Gets the current status of the InvokeOCR_Watcher Scheduled Task.

.DESCRIPTION
    Queries the Windows Task Scheduler for the InvokeOCR_Watcher task and displays
    its current state with color-coded output:
    - RUNNING (Green)  : The watcher is actively monitoring C:\scans for new files.
    - READY (Yellow)   : The task is registered but not currently running.
    - DISABLED (Gray)  : The task has been manually disabled in Task Scheduler.
    - NOT INSTALLED (Red) : The task is not registered. Run Install-OCRWatcher.ps1.

.EXAMPLE
    .\Get-OCRWatcherStatus.ps1

    Displays the current watcher status.

.NOTES
    Does not require Administrator privileges.

    See also:
    - Install-OCRWatcher.ps1 - Install the watcher
    - Remove-OCRWatcher.ps1  - Uninstall the watcher
#>

$taskName = "InvokeOCR_Watcher"
$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

if (-not $task) {
    Write-Host "Service Status: " -NoNewline
    Write-Host "NOT INSTALLED" -ForegroundColor Red
    Write-Host "The background watcher is not currently registered on this system."
} else {
    Write-Host "Service Status: " -NoNewline
    if ($task.State -eq 'Running') {
        Write-Host "RUNNING" -ForegroundColor Green
        Write-Host "The folder watcher is actively listening to C:\scans."
    } elseif ($task.State -eq 'Ready') {
        Write-Host "READY (Idling / Stopped)" -ForegroundColor Yellow
        Write-Host "The task is installed but is not currently running. You can start it via Task Scheduler or by re-running the installer."
    } elseif ($task.State -eq 'Disabled') {
        Write-Host "DISABLED" -ForegroundColor Gray
        Write-Host "The task is installed but has been manually disabled."
    } else {
        Write-Host $task.State -ForegroundColor Cyan
    }
}
