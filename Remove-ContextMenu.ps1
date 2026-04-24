<#
.SYNOPSIS
    Removes the "Make Searchable (Invoke-OCR)" option from the Windows Right-Click Context Menu.
#>

# Auto-elevate to Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Administrator privileges required. Requesting elevation..." -ForegroundColor Yellow
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$Extensions = @(".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")

# Map HKCR: drive (not available by default in PowerShell)
if (-not (Get-PSDrive -Name HKCR -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
}

Write-Host "Removing context menu handlers..."
foreach ($ext in $Extensions) {
    $KeyPath = "HKCR:\SystemFileAssociations\$ext\shell\InvokeOCR"
    if (Test-Path $KeyPath) {
        Remove-Item -Path $KeyPath -Recurse -Force
    }
}

Write-Host "Context menu removed successfully!" -ForegroundColor Green
Start-Sleep -Seconds 3
