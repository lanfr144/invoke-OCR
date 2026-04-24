<#
.SYNOPSIS
    Adds the "Make Searchable (Invoke-OCR)" option to the Windows Right-Click Context Menu.
#>

# Auto-elevate to Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Administrator privileges required. Requesting elevation..." -ForegroundColor Yellow
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$ScriptPath = Join-Path $PSScriptRoot "Invoke-OCR.ps1"
$Extensions = @(".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")

# Map HKCR: drive (not available by default in PowerShell)
if (-not (Get-PSDrive -Name HKCR -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
}

Write-Host "Installing context menu handlers..."
foreach ($ext in $Extensions) {
    $KeyPath = "HKCR:\SystemFileAssociations\$ext\shell\InvokeOCR"
    $CommandPath = "$KeyPath\command"
    
    if (-not (Test-Path $KeyPath)) { New-Item -Path $KeyPath -Force | Out-Null }
    if (-not (Test-Path $CommandPath)) { New-Item -Path $CommandPath -Force | Out-Null }
    
    Set-ItemProperty -Path $KeyPath -Name "(default)" -Value "Make Searchable (Invoke-OCR)"
    Set-ItemProperty -Path $KeyPath -Name "Icon" -Value "imageres.dll,-5313" 
    
    $runCmd = "`"powershell.exe`" -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"& '$ScriptPath' -Path '%1' -Silent -y`""
    Set-ItemProperty -Path $CommandPath -Name "(default)" -Value $runCmd
}

Write-Host "Context menu installed successfully! You can now right-click PDFs and Images." -ForegroundColor Green
Start-Sleep -Seconds 3
