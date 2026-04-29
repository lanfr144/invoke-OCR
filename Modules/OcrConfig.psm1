# Requires OcrLogging.psm1 to be imported

# Valid config keys for validation
$script:validConfigKeys = @(
    "Language", "Dpi", "Page", "WatermarkPdf", "ThrottleLimit",
    "ForceOCR", "RemoveSource",
    "TesseractPath", "GhostscriptPath", "PdftkPath",
    "TesseractArgs", "GhostscriptArgs", "PdftkArgs",
    "MoveSourceDir", "MoveOcrDir", "MoveTxtDir",
    "EmailTo", "EmailFiles", "EmailSubject", "EmailBody",
    "EmailFrom", "EmailReplyTo",
    "SmtpServer", "SmtpPort", "SmtpUser", "SmtpPassword"
)

function Get-ValidConfigKeys {
    return $script:validConfigKeys
}

function Import-OcrConfigFile {
    param([string]$FilePath)
    $config = @{}
    if (-not (Test-Path -LiteralPath $FilePath)) { return $config }

    $lines = Get-Content -LiteralPath $FilePath
    $lineNum = 0
    foreach ($line in $lines) {
        $lineNum++
        $line = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line -match "^(#|')") { continue }

        $idx = $line.IndexOf('=')
        if ($idx -gt 0) {
            $key = $line.Substring(0, $idx).Trim()
            $value = $line.Substring($idx + 1).Trim()

            # Strip surrounding quotes if present
            if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
                $value = $value.Substring(1, $value.Length - 2)
            }

            # Validate key
            if ($key -notin $script:validConfigKeys) {
                Write-Warn "Unknown config key '$key' at line $lineNum in $FilePath (did you mean one of: $($script:validConfigKeys -join ', '))"
            }

            $config[$key] = $value
        }
    }
    return $config
}

function Import-OcrConfig {
    param([string]$ConfigPath)

    # Hierarchical config: walk up from file directory to drive root, collecting configs
    $configDir = Split-Path -Parent $ConfigPath
    $configStack = @()

    $currentDir = $configDir
    while ($currentDir -and (Test-Path $currentDir)) {
        $cfgFile = Join-Path $currentDir ".ocrconfig"
        if (Test-Path -LiteralPath $cfgFile) {
            $configStack += $cfgFile
        }
        $parentDir = Split-Path -Parent $currentDir
        if ($parentDir -eq $currentDir) { break }  # reached root
        $currentDir = $parentDir
    }

    # Merge from root to leaf (child overrides parent)
    $merged = @{}
    for ($i = $configStack.Count - 1; $i -ge 0; $i--) {
        $cfg = Import-OcrConfigFile $configStack[$i]
        foreach ($key in $cfg.Keys) {
            $merged[$key] = $cfg[$key]
        }
    }

    if ($merged.Count -gt 0) {
        Write-Info "Loaded .ocrconfig ($($merged.Count) settings from $($configStack.Count) file(s))"
    }

    # Handle secure password via Windows Credential Manager
    if ($merged.ContainsKey("SmtpPassword") -and $merged["SmtpPassword"] -match "^credential:(.+)$") {
        $credName = $Matches[1]
        try {
            $credPath = Join-Path $env:USERPROFILE ".ocrCredentials_$credName.xml"
            if (Test-Path -LiteralPath $credPath) {
                $cred = Import-Clixml -LiteralPath $credPath
                $merged["SmtpUser"] = $cred.UserName
                $merged["SmtpPassword"] = $cred.GetNetworkCredential().Password
                Write-Info "Loaded SMTP credentials from secure store '$credName'"
            }
            else {
                Write-Warn "Credential file not found: $credPath. Use Save-OcrCredential.ps1 to create it."
            }
        }
        catch {
            Write-Warn "Failed to load credential '$credName': $_"
        }
    }

    return $merged
}

Export-ModuleMember -Function Get-ValidConfigKeys, Import-OcrConfigFile, Import-OcrConfig
