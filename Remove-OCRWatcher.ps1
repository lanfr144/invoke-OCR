<#
.SYNOPSIS
    Removes the InvokeOCR_Watcher Scheduled Task from the system.
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
