<#
.SYNOPSIS
    Removes the InvokeOCR_Watcher Scheduled Task from the system.
#>

$taskName = "InvokeOCR_Watcher"
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
