<#
.SYNOPSIS
    Background watcher script that monitors C:\scans and its subdirectories to automatically trigger Invoke-OCR.ps1.

.DESCRIPTION
    This script is designed to be run as a Scheduled Task. It uses System.IO.FileSystemWatcher to detect when 
    new PDFs or images are dropped into C:\scans. Based on the subfolder (en, de, fr, lb), it maps to the correct 
    Tesseract language and triggers Invoke-OCR.ps1.
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

    # Determine Language based on folder
    $lang = "eng+fra+deu+ltz+por+lat" # default fallback
    $dirName = (Split-Path -Leaf $dir).ToLower()

    switch ($dirName) {
        "en" { $lang = "eng" }
        "de" { $lang = "deu" }
        "fr" { $lang = "fra" }
        "lb" { $lang = "ltz" }
    }

    Write-Host "Triggering Invoke-OCR for $path with Language $lang..."
    
    # Trigger the script silently
    $argsList = @(
        "-File", "`"$InvokeScript`"",
        "-Path", "`"$path`"",
        "-Language", "`"$lang`"",
        "-Silent",
        "-y"
    )
    
    # Start the process hidden so it doesn't interrupt the user
    Start-Process -FilePath "pwsh" -ArgumentList $argsList -WindowStyle Hidden -Wait
}

# Register the event subscriber
Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $action -SourceIdentifier "OCRWatcher_Created"

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
