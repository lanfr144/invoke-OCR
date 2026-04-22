<#
.SYNOPSIS
    Gets the current status of the InvokeOCR_Watcher Scheduled Task.
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
