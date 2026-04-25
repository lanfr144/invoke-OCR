<#
.SYNOPSIS
    Adds the "Make Searchable (Invoke-OCR)" option to the Windows Right-Click Context Menu.

.DESCRIPTION
    Registers a Windows Explorer context menu entry for supported file types that allows
    right-clicking any PDF or image file and selecting "Make Searchable (Invoke-OCR)" to
    trigger OCR processing directly from the file explorer.

    Supported file extensions: .pdf, .png, .jpg, .jpeg, .bmp, .tif, .tiff

    The context menu entry is added to the HKEY_CLASSES_ROOT registry under
    SystemFileAssociations for each supported extension. The OCR runs hidden in the
    background using PowerShell with -Silent and -y flags.

    Requires Administrator privileges (auto-elevates if needed).
    To remove the context menu, run Remove-ContextMenu.ps1.

.EXAMPLE
    .\Install-ContextMenu.ps1

    Installs the right-click menu. Will auto-request Administrator elevation if needed.

.NOTES
    Requires: Administrator privileges (auto-elevates)
    
    The context menu runs Invoke-OCR.ps1 with -Silent -y flags for unattended processing.
    If a .ocrconfig file exists in the same directory as the file, it will be used.

    See also:
    - Remove-ContextMenu.ps1 - Remove the context menu entry
    - Invoke-OCR.ps1         - The OCR processing script
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
