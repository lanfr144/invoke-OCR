<#
.SYNOPSIS
    Securely saves SMTP credentials for use in .ocrconfig files.

.DESCRIPTION
    Prompts for SMTP username and password, then encrypts and saves them using
    Windows DPAPI (Data Protection API) via Export-Clixml. The credentials are
    encrypted with the current user's Windows account and can only be decrypted
    by the same user on the same machine.

    To use saved credentials in .ocrconfig, set:
        SmtpPassword = credential:<name>

    Where <name> matches the -Name parameter used when saving.

.PARAMETER Name
    A label for the credential. Used to identify different SMTP accounts.
    Default: "default"

.EXAMPLE
    .\Save-OcrCredential.ps1

    Saves credentials with the label "default".

.EXAMPLE
    .\Save-OcrCredential.ps1 -Name "office365"

    Saves credentials with the label "office365".
    In .ocrconfig use: SmtpPassword = credential:office365

.NOTES
    The credential file is saved to: %USERPROFILE%\.ocrCredentials_<name>.xml
    It is encrypted with Windows DPAPI and cannot be read by other users or machines.

    See also:
    - Invoke-OCR.ps1 - Uses these credentials when SmtpPassword = credential:<name>
#>
param(
    [string]$Name = "default"
)

$credPath = Join-Path $env:USERPROFILE ".ocrCredentials_$Name.xml"

Write-Host "=== Save OCR SMTP Credential ===" -ForegroundColor Cyan
Write-Host "This will securely store SMTP credentials encrypted with your Windows account."
Write-Host "Credential file: $credPath"
Write-Host ""

$credential = Get-Credential -Message "Enter SMTP username and password for credential '$Name'"

if (-not $credential) {
    Write-Host "No credentials provided. Cancelled." -ForegroundColor Yellow
    exit
}

$credential | Export-Clixml -LiteralPath $credPath -Force
Write-Host "`nCredentials saved successfully!" -ForegroundColor Green
Write-Host "To use in .ocrconfig, add:"
Write-Host "    SmtpPassword = credential:$Name" -ForegroundColor Cyan
Write-Host "`nThe credential is encrypted with your Windows account (DPAPI) and cannot be read by other users."
