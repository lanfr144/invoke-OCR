<#
.SYNOPSIS
    Background watcher script that monitors C:\scans and its subdirectories to automatically trigger Invoke-OCR.ps1.

.DESCRIPTION
    This script is designed to be run as a Scheduled Task (registered by Install-OCRWatcher.ps1).
    It uses System.IO.FileSystemWatcher to detect when new PDFs or images are dropped into C:\scans.
    
    Language detection priority:
    1. Base default: eng+fra+deu+ltz+por+lat (all languages)
    2. Directory-based override: en->eng, de->deu, fr->fra, lb->ltz
    3. Per-directory .ocrconfig file (highest priority)

    The .ocrconfig file supports all Invoke-OCR.ps1 parameters in Key=Value format.
    See Get-Help Invoke-OCR.ps1 -Full for the complete list of supported config keys.

    Supported file types: .pdf, .png, .jpg, .jpeg, .tif, .tiff, .bmp

    Safety features:
    - Waits up to 60 seconds for file locks to be released (scanner/OS writing)
    - Ignores _ocr.pdf, _ocr.txt, and .err.log output files to prevent loops
    - Gracefully cleans up FileSystemWatcher on Ctrl+C or termination

.EXAMPLE
    .\Start-OCRWatcher.ps1

    Starts monitoring C:\scans. Usually run automatically by the Scheduled Task.

.NOTES
    This script is not intended to be run directly by users. Use Install-OCRWatcher.ps1
    to register it as a background Scheduled Task.

    See also:
    - Install-OCRWatcher.ps1   - Register this script as a service
    - Get-OCRWatcherStatus.ps1 - Check if the watcher is running
    - Remove-OCRWatcher.ps1    - Uninstall the watcher

.LINK
    https://learn.microsoft.com/en-us/dotnet/api/system.io.filesystemwatcher
#>

$WatchFolder = "C:\scans"
$InvokeScript = Join-Path $PSScriptRoot "Invoke-OCR.ps1"

if (-not (Test-Path $WatchFolder)) {
    Write-Warning "Watch folder $WatchFolder does not exist. Waiting for it to be created..."
    while (-not (Test-Path $WatchFolder)) { Start-Sleep -Seconds 10 }
}

# Create a FileSystemWatcher
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $WatchFolder
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true

# Define the action to take when a file is created
$action = {
    $InvokeScript = $Event.MessageData
    $path = $Event.SourceEventArgs.FullPath
    $name = $Event.SourceEventArgs.Name
    $changeType = $Event.SourceEventArgs.ChangeType
    $dir = Split-Path -Parent $path
    
    # Ignore the _ocr output files so we don't loop endlessly
    if ($name -match "_ocr\.(pdf|txt)$" -or $name -match "\.err\.log$") {
        return
    }

    # Only process supported extensions
    $ext = [System.IO.Path]::GetExtension($path).ToLower()
    if ($ext -notmatch "^\.(pdf|png|jpg|jpeg|tif|tiff|bmp)$") {
        return
    }

    Write-Host "Detected new file: $path"

    # Wait for file to be released by the OS/Scanner
    $fileLocked = $true
    $attempts = 0
    while ($fileLocked -and $attempts -lt 30) {
        try {
            $stream = [System.IO.File]::Open($path, 'Open', 'Read', 'None')
            $stream.Close()
            $stream.Dispose()
            $fileLocked = $false
        } catch {
            Start-Sleep -Seconds 2
            $attempts++
        }
    }

    if ($fileLocked) {
        Write-Warning "Timed out waiting for file lock to release on $path"
        return
    }

    # 1. Base default fallback
    $lang = "eng+fra+deu+ltz+por+lat" 
    $dirName = (Split-Path -Leaf $dir).ToLower()

    # 2. Directory based override
    switch ($dirName) {
        "en" { $lang = "eng" }
        "de" { $lang = "deu" }
        "fr" { $lang = "fra" }
        "lb" { $lang = "ltz" }
    }

    # Base Arguments
    $argsList = @(
        "-File", "`"$InvokeScript`"",
        "-Path", "`"$path`"",
        "-Silent",
        "-y"
    )

    $customLang = $null

    # 3. Config file based override
    $configPath = Join-Path $dir ".ocrconfig"
    if (Test-Path -LiteralPath $configPath) {
        Write-Host "Found .ocrconfig file in $dir"
        $lines = Get-Content -LiteralPath $configPath
        foreach ($line in $lines) {
            $line = $line.Trim()
            # Ignore comments and empty lines
            if ([string]::IsNullOrWhiteSpace($line) -or $line -match "^(#|')") { continue }
            
            # Parse Key=Value
            $idx = $line.IndexOf('=')
            if ($idx -gt 0) {
                $key = $line.Substring(0, $idx).Trim()
                $value = $line.Substring($idx + 1).Trim()
                
                # Strip surrounding quotes if present
                if (($value.StartsWith("`"") -and $value.EndsWith("`"")) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
                    $value = $value.Substring(1, $value.Length - 2)
                }

                if ($key -ieq "Language") {
                    $customLang = $value
                } else {
                    # Inject arbitrary parameter
                    $argsList += "-$key"
                    if (-not [string]::IsNullOrWhiteSpace($value)) {
                        $argsList += "`"$value`""
                    }
                }
            }
        }
    }

    if ($customLang) {
        $lang = $customLang
    }
    
    $argsList += "-Language"
    $argsList += "`"$lang`""

    Write-Host "Triggering Invoke-OCR for $path with Language $lang..."
    
    # Start the process hidden so it doesn't interrupt the user
    Start-Process -FilePath "pwsh" -ArgumentList $argsList -WindowStyle Hidden -Wait
}

# Register the event subscriber
Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $action -SourceIdentifier "OCRWatcher_Created" -MessageData $InvokeScript

Write-Host "Monitoring $WatchFolder and subdirectories for new files. Press Ctrl+C to stop..."

# Keep the script running infinitely
try {
    while ($true) {
        Start-Sleep -Seconds 5
    }
} finally {
    Unregister-Event -SourceIdentifier "OCRWatcher_Created" -ErrorAction SilentlyContinue
    $watcher.Dispose()
}
