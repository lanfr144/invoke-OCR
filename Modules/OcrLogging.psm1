$script:Quiet = $false
$script:Silent = $false

function Set-OcrLoggingState {
    param([switch]$Quiet, [switch]$Silent)
    $script:Quiet = $Quiet
    $script:Silent = $Silent
}

function Write-Info {
    param([string]$Message)
    if (-not $script:Quiet -and -not $script:Silent) { Write-Host $Message }
}

function Write-Warn {
    param([string]$Message)
    if (-not $script:Silent) { Write-Warning $Message }
}

function Write-Err {
    param([string]$Message)
    if (-not $script:Silent) { Write-Error -Message $Message }
}

function Write-SystemLog {
    param(
        [string]$Message, 
        [string]$Type = "Information",
        [string]$LogDirectory = $PSScriptRoot
    )
    # Using parent directory since module is in Modules/
    $baseDir = Split-Path -Parent $PSScriptRoot
    if (-not [string]::IsNullOrEmpty($LogDirectory)) {
        $baseDir = $LogDirectory
    } else {
        $baseDir = Split-Path -Parent $PSScriptRoot
    }
    $logPath = Join-Path $baseDir "ocr_service.log"

    # Log rotation: rename to .log.bak if exceeds 5MB
    if (Test-Path -LiteralPath $logPath) {
        $logSize = (Get-Item -LiteralPath $logPath).Length
        if ($logSize -gt 5MB) {
            $bakPath = "$logPath.bak"
            Move-Item -LiteralPath $logPath -Destination $bakPath -Force -ErrorAction SilentlyContinue
        }
    }

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$timestamp] [$Type] $Message"
    Add-Content -LiteralPath $logPath -Value $line -ErrorAction SilentlyContinue

    try {
        if (-not [System.Diagnostics.EventLog]::SourceExists("Invoke-OCR")) {
            [System.Diagnostics.EventLog]::CreateEventSource("Invoke-OCR", "Application")
        }
        [System.Diagnostics.EventLog]::WriteEntry("Invoke-OCR", $Message, $Type, 1)
    }
    catch { }
}

Export-ModuleMember -Function Set-OcrLoggingState, Write-Info, Write-Warn, Write-Err, Write-SystemLog
