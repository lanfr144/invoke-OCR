<#
.SYNOPSIS
    Validates .ocrconfig files for correctness against Invoke-OCR.ps1 parameters.

.DESCRIPTION
    Scans one or more directories for .ocrconfig files and validates that all keys
    are recognized Invoke-OCR.ps1 parameters. Reports valid keys, unknown keys,
    and hierarchical parent configs that will be inherited.

    This script can validate:
    - A single directory
    - Multiple directories
    - An entire tree (recursive mode)

    It uses the same key validation list as Invoke-OCR.ps1 to ensure consistency.

.PARAMETER Path
    One or more directory paths to check for .ocrconfig files.
    Accepts pipeline input.

.PARAMETER Recurse
    When set, searches subdirectories recursively for all .ocrconfig files.

.EXAMPLE
    .\Test-OcrConfig.ps1 -Path "C:\scans"

    Validates the .ocrconfig file in C:\scans.

.EXAMPLE
    .\Test-OcrConfig.ps1 -Path "C:\scans" -Recurse

    Recursively finds and validates all .ocrconfig files under C:\scans.

.EXAMPLE
    "C:\scans\en", "C:\scans\fr" | .\Test-OcrConfig.ps1

    Validates configs in multiple directories via pipeline.

.NOTES
    See also:
    - Invoke-OCR.ps1 -ValidateConfig  - Inline validation mode
    - Save-OcrCredential.ps1          - Secure credential storage

.LINK
    https://github.com/lanfr144/invoke-OCR
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias("FullName")]
    [string[]]$Path,

    [switch]$Recurse
)

BEGIN {
    # Valid config keys (must match Invoke-OCR.ps1)
    $validKeys = @(
        "Language", "Dpi", "Page", "WatermarkPdf", "ThrottleLimit",
        "ForceOCR", "RemoveSource",
        "TesseractPath", "GhostscriptPath", "PdftkPath",
        "TesseractArgs", "GhostscriptArgs", "PdftkArgs",
        "MoveSourceDir", "MoveOcrDir", "MoveTxtDir",
        "EmailTo", "EmailFiles", "EmailSubject", "EmailBody",
        "EmailFrom", "EmailReplyTo",
        "SmtpServer", "SmtpPort", "SmtpUser", "SmtpPassword"
    )

    $totalFiles = 0
    $totalErrors = 0

    function Test-ConfigFile {
        param([string]$FilePath)

        if (-not (Test-Path -LiteralPath $FilePath)) { return }

        $script:totalFiles++
        Write-Host "`n=== Validating: $FilePath ===" -ForegroundColor Cyan

        $lines = Get-Content -LiteralPath $FilePath
        $lineNum = 0
        $keys = @{}
        $errors = 0

        foreach ($line in $lines) {
            $lineNum++
            $trimmed = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed -match "^(#|')") { continue }

            $idx = $trimmed.IndexOf('=')
            if ($idx -le 0) {
                Write-Host "  [!!] Line ${lineNum}: Invalid syntax (no '=' found): $trimmed" -ForegroundColor Red
                $errors++
                continue
            }

            $key = $trimmed.Substring(0, $idx).Trim()
            $value = $trimmed.Substring($idx + 1).Trim()

            # Strip quotes
            if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
                $value = $value.Substring(1, $value.Length - 2)
            }

            if ($key -in $validKeys) {
                Write-Host "  [OK] Line ${lineNum}: $key = $value" -ForegroundColor Green
            }
            else {
                Write-Host "  [!!] Line ${lineNum}: Unknown key '$key' = $value" -ForegroundColor Red
                
                # Suggest closest match
                $closest = $validKeys | Where-Object { $_ -like "*$key*" -or $key -like "*$_*" } | Select-Object -First 1
                if ($closest) {
                    Write-Host "       Did you mean: $closest ?" -ForegroundColor Yellow
                }
                $errors++
            }

            # Duplicate key check
            if ($keys.ContainsKey($key)) {
                Write-Host "  [!!] Line ${lineNum}: Duplicate key '$key' (overrides line $($keys[$key]))" -ForegroundColor Yellow
            }
            $keys[$key] = $lineNum
        }

        # Check for credential references
        if ($keys.ContainsKey("SmtpPassword")) {
            $pwLine = ($lines | Select-Object -Skip ($keys["SmtpPassword"] - 1) | Select-Object -First 1).Trim()
            $pwIdx = $pwLine.IndexOf('=')
            $pwVal = $pwLine.Substring($pwIdx + 1).Trim().Trim('"', "'")
            if ($pwVal -match "^credential:(.+)$") {
                $credName = $Matches[1]
                $credPath = Join-Path $env:USERPROFILE ".ocrCredentials_$credName.xml"
                if (Test-Path -LiteralPath $credPath) {
                    Write-Host "  [OK] Credential '$credName' exists at $credPath" -ForegroundColor Green
                }
                else {
                    Write-Host "  [!!] Credential '$credName' NOT found. Run Save-OcrCredential.ps1 -Name '$credName'" -ForegroundColor Red
                    $errors++
                }
            }
        }

        # Check parent config inheritance
        $parentDir = Split-Path -Parent (Split-Path -Parent $FilePath)
        if ($parentDir) {
            $parentConfig = Join-Path $parentDir ".ocrconfig"
            if (Test-Path -LiteralPath $parentConfig) {
                Write-Host "  [INFO] Parent config: $parentConfig (values will be inherited)" -ForegroundColor DarkCyan
            }
        }

        if ($errors -eq 0) {
            Write-Host "  Result: VALID ($($keys.Count) keys)" -ForegroundColor Green
        }
        else {
            Write-Host "  Result: $errors error(s) found" -ForegroundColor Red
            $script:totalErrors += $errors
        }
    }
}

PROCESS {
    foreach ($p in $Path) {
        if (-not (Test-Path $p)) {
            Write-Host "Path not found: $p" -ForegroundColor Red
            continue
        }

        $item = Get-Item -LiteralPath $p
        if (-not $item.PSIsContainer) {
            # Direct file path
            Test-ConfigFile $item.FullName
            continue
        }

        # Directory: look for .ocrconfig
        if ($Recurse) {
            $configs = Get-ChildItem -Path $item.FullName -Filter ".ocrconfig" -Recurse -File -Force -ErrorAction SilentlyContinue
            if ($configs.Count -eq 0) {
                Write-Host "No .ocrconfig files found under $($item.FullName)" -ForegroundColor Yellow
            }
            foreach ($cfg in $configs) {
                Test-ConfigFile $cfg.FullName
            }
        }
        else {
            $configFile = Join-Path $item.FullName ".ocrconfig"
            if (Test-Path -LiteralPath $configFile) {
                Test-ConfigFile $configFile
            }
            else {
                Write-Host "No .ocrconfig found in $($item.FullName)" -ForegroundColor Yellow
            }
        }
    }
}

END {
    Write-Host "`n--- Summary ---" -ForegroundColor Cyan
    Write-Host "Files checked: $totalFiles"
    if ($totalErrors -eq 0) {
        Write-Host "All configs are valid!" -ForegroundColor Green
    }
    else {
        Write-Host "Total errors: $totalErrors" -ForegroundColor Red
    }
}
