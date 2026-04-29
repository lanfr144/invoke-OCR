function Invoke-OcrTask {
    [CmdletBinding()]
param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias("FullName")]
    [string[]]$Path,

    # Supports one or many Tesseract language codes, array format.
    [string[]]$Language = @("eng", "fra", "deu", "ltz", "por", "lat"),

    # DPI for PDF bursting via Ghostscript and Tesseract OCR
    [int]$Dpi = 300,

    # Specific page to process. If 0 or omitted, processes all pages.
    [int]$Page = 0,

    # Path to a PDF file to use as a watermark (philigram).
    [string]$WatermarkPdf,

    # Number of pages to process concurrently in Tesseract. Requires PowerShell 7.
    [int]$ThrottleLimit = 4,

    # Bypasses the confirmation prompt asking if the script found the correct executables.
    [Alias("y")]
    [switch]$Yes,

    # Suppresses standard output text (like "Extracting pages..."), but still prints errors/warnings.
    [switch]$Quiet,

    # Suppresses EVERYTHING. No success messages, no warnings, no errors to the console.
    [switch]$Silent,

    # If the PDF already has a text layer, process it anyway (good for documents with schemas/images missing OCR)
    [switch]$ForceOCR,

    # Custom Explicit Software Paths (Overrides the automatic search)
    [string]$TesseractPath,
    [string]$GhostscriptPath,
    [string]$PdftkPath,

    # Additional custom parameters to pass verbatim into Tesseract
    [string[]]$TesseractArgs,

    # Additional custom parameters to pass verbatim into Ghostscript
    [string[]]$GhostscriptArgs,

    # Additional custom parameters to pass verbatim into PDFtk
    [string[]]$PdftkArgs,

    # File Movement
    [string]$MoveSourceDir,
    [string]$MoveOcrDir,
    [string]$MoveTxtDir,
    [switch]$RemoveSource,

    # Emailing
    [string[]]$EmailTo,
    [string[]]$EmailFiles = @("Ocr"),
    [string]$EmailSubject = 'OCR Completed: ${filename}',
    [string]$EmailBody = 'The document ${filename} was successfully processed in ${elapsed} seconds.',
    [string]$EmailFrom = 'Invoke-OCR <no-reply@localhost>',
    [string]$EmailReplyTo,
    [string]$SmtpServer,
    [int]$SmtpPort = 25,
    [string]$SmtpUser,
    [string]$SmtpPassword,

    # Validation mode: only validate .ocrconfig files, do not process
    [switch]$ValidateConfig,

    # Path to save the CSV report
    [string]$ReportPath,

    # Skip extracting and injecting PDF metadata
    [switch]$SkipMetadataExtract,

    # Skip extracting PDF form fields
    [switch]$SkipFieldsExtract
)

BEGIN {
    function Verify-Executable {
        param([string]$ExePath, [string]$TestArg, [string]$RegexMatch)
        try {
            $output = & $ExePath $TestArg 2>&1
            $outStr = $output -join "`n"
            if ($outStr -match $RegexMatch) { return $true }
        }
        catch { }
        return $false
    }
    
    function Confirm-Executable {
        param([string]$CandidateUrl)
        if ($Yes) { return $true }
        
        while ($true) {
            # Bypasses host UI issues in some older terminals to strictly fetch 'y' or 'n'
            $ans = Read-Host "Found executable at: $CandidateUrl `nIs this correct? (y/n)"
            if ($ans -match "^y") { return $true }
            if ($ans -match "^n") { return $false }
            if ([string]::IsNullOrWhiteSpace($ans)) { return $true }
        }
    }

    function Find-Executable {
        param(
            [string]$ExplicitPath,
            [string]$ExePattern,
            [string]$FolderPattern,
            [string]$VerifyArg,
            [string]$VerifyRegex
        )

        # 1. Provided as arguments
        if ($ExplicitPath -and (Test-Path $ExplicitPath -PathType Leaf)) {
            if (Verify-Executable $ExplicitPath $VerifyArg $VerifyRegex) { 
                if (Confirm-Executable $ExplicitPath) { return $ExplicitPath }
            }
            else { Write-Warn "Explicit path $ExplicitPath did not pass validation check." }
        }

        # 2. Within system PATH variable
        $cmds = Get-Command $ExePattern -ErrorAction SilentlyContinue | Where-Object CommandType -eq 'Application'
        foreach ($cmd in $cmds) {
            $candidate = $cmd.Source
            if ($candidate -match 'uninstall|unins') { continue }
            if (Verify-Executable $candidate $VerifyArg $VerifyRegex) { 
                if (Confirm-Executable $candidate) { return $candidate }
            }
        }

        # 3. Default well known places
        $baseDirs = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Select-Object -Unique | Where-Object { $_ -ne $null }
        foreach ($baseDir in $baseDirs) {
            if (-not (Test-Path $baseDir)) { continue }
            $folders = Get-ChildItem -Path $baseDir -Filter $FolderPattern -Directory -ErrorAction SilentlyContinue
            foreach ($folder in $folders) {
                $exes = Get-ChildItem -Path $folder.FullName -Filter $ExePattern -Recurse -File -ErrorAction SilentlyContinue
                foreach ($exe in $exes) {
                    if ($exe.FullName -match 'uninstall|unins') { continue }
                    if (Verify-Executable $exe.FullName $VerifyArg $VerifyRegex) { 
                        if (Confirm-Executable $exe.FullName) { return $exe.FullName }
                    }
                }
            }
        }
        return $null
    }

    Write-Info "Resolving executables..."
    $tesseract = Find-Executable -ExplicitPath $TesseractPath -ExePattern "tesser*.exe" -FolderPattern "tesser*" -VerifyArg "-v" -VerifyRegex "(?i)tesseract"
    $ghostscript = Find-Executable -ExplicitPath $GhostscriptPath -ExePattern "gswin*c.exe" -FolderPattern "gs*" -VerifyArg "-v" -VerifyRegex "(?i)Ghostscript"
    if (-not $ghostscript) {
        $ghostscript = Find-Executable -ExplicitPath $GhostscriptPath -ExePattern "gswin*.exe" -FolderPattern "gs*" -VerifyArg "-v" -VerifyRegex "(?i)Ghostscript"
    }
    $pdftk = Find-Executable -ExplicitPath $PdftkPath -ExePattern "pdftk*.exe" -FolderPattern "pdftk*" -VerifyArg "--help" -VerifyRegex "(?i)pdftk"

    $halt = $false
    if (-not $tesseract) { Write-Err "Tesseract could not be found via args, PATH, or default directories."; $halt = $true }
    if (-not $ghostscript) { Write-Err "Ghostscript could not be found via args, PATH, or default directories."; $halt = $true }
    if (-not $pdftk) { Write-Err "PDFtk could not be found via args, PATH, or default directories."; $halt = $true }
    
    if (-not $ValidateConfig) {
        if ($halt) { throw "Missing required prerequisites." }
        Write-Info "Ready. All executables verified."
    }
}

PROCESS {
    $reportData = @()
    # ValidateConfig mode: validate .ocrconfig files and exit
    if ($ValidateConfig) {
        foreach ($p in $Path) {
            if ([string]::IsNullOrWhiteSpace($p)) { continue }
            try {
                $item = Get-Item -LiteralPath $p -ErrorAction Stop
                $dir = if ($item.PSIsContainer) { $item.FullName } else { $item.DirectoryName }
            }
            catch {
                Write-Err "Path not found: $p"
                continue
            }

            $configFile = Join-Path $dir ".ocrconfig"
            if (-not (Test-Path -LiteralPath $configFile)) {
                Write-Host "No .ocrconfig found in $dir" -ForegroundColor Yellow
                continue
            }

            Write-Host "`n=== Validating: $configFile ===" -ForegroundColor Cyan
            $config = Import-OcrConfigFile $configFile
            $errorCount = 0

            foreach ($key in $config.Keys) {
                if ($key -in (Get-ValidConfigKeys)) {
                    Write-Host "  [OK] $key = $($config[$key])" -ForegroundColor Green
                }
                else {
                    Write-Host "  [!!] Unknown key: '$key' = $($config[$key])" -ForegroundColor Red
                    $errorCount++
                }
            }

            # Also check for hierarchical parent configs
            $parentDir = Split-Path -Parent $dir
            if ($parentDir -and $parentDir -ne $dir) {
                $parentConfig = Join-Path $parentDir ".ocrconfig"
                if (Test-Path -LiteralPath $parentConfig) {
                    Write-Host "  [INFO] Parent config found: $parentConfig (values will be inherited)" -ForegroundColor DarkCyan
                }
            }

            if ($errorCount -eq 0) {
                Write-Host "  Result: VALID ($($config.Count) keys)" -ForegroundColor Green
            }
            else {
                Write-Host "  Result: $errorCount INVALID key(s) found" -ForegroundColor Red
            }
        }
        return
    }

    $langStr = $Language -join "+"

    foreach ($p in $Path) {
        $startTime = Get-Date
        if ([string]::IsNullOrWhiteSpace($p)) { continue }

        try {
            $file = Get-Item -LiteralPath $p -ErrorAction Stop
        }
        catch {
            Write-Err "File not found: $p"
            continue
        }

        $ext = $file.Extension.ToLower()
        $outPath = Join-Path $file.DirectoryName ($file.BaseName + "_ocr.pdf")
        $errPath = Join-Path $file.DirectoryName ($file.BaseName + ".err.log")

        # Load per-directory .ocrconfig (CLI parameters take precedence)
        $configPath = Join-Path $file.DirectoryName ".ocrconfig"
        $ocrConfig = Import-OcrConfig $configPath

        # Map for simple string/int parameters: ConfigKey -> VariableName
        $configMap = @{
            "Language"     = "Language"
            "Dpi"          = "Dpi"
            "WatermarkPdf" = "WatermarkPdf"
            "MoveSourceDir"= "MoveSourceDir"
            "MoveOcrDir"   = "MoveOcrDir"
            "MoveTxtDir"   = "MoveTxtDir"
            "EmailTo"      = "EmailTo"
            "EmailSubject" = "EmailSubject"
            "EmailBody"    = "EmailBody"
            "EmailFrom"    = "EmailFrom"
            "EmailReplyTo" = "EmailReplyTo"
            "SmtpServer"   = "SmtpServer"
            "SmtpPort"     = "SmtpPort"
            "SmtpUser"     = "SmtpUser"
            "SmtpPassword" = "SmtpPassword"
            "SkipMetadataExtract" = "SkipMetadataExtract"
            "SkipFieldsExtract"   = "SkipFieldsExtract"
        }

        foreach ($cfgKey in $configMap.Keys) {
            $varName = $configMap[$cfgKey]
            if ($ocrConfig.ContainsKey($cfgKey) -and -not $PSBoundParameters.ContainsKey($varName)) {
                $cfgValue = $ocrConfig[$cfgKey]
                switch ($varName) {
                    "Language"  { $Language = $cfgValue -split '\+'; $langStr = $cfgValue }
                    "Dpi"       { $Dpi = [int]$cfgValue }
                    "SmtpPort"  { $SmtpPort = [int]$cfgValue }
                    "EmailTo"   { $EmailTo = @($cfgValue) }
                    default     { Set-Variable -Name $varName -Value $cfgValue }
                }
            }
        }

        # Handle switch parameters from config
        if ($ocrConfig.ContainsKey("ForceOCR") -and -not $PSBoundParameters.ContainsKey("ForceOCR")) {
            if ($ocrConfig["ForceOCR"] -match "^(true|1|yes)$") { $ForceOCR = $true }
        }
        if ($ocrConfig.ContainsKey("RemoveSource") -and -not $PSBoundParameters.ContainsKey("RemoveSource")) {
            if ($ocrConfig["RemoveSource"] -match "^(true|1|yes)$") { $RemoveSource = $true }
        }
        if ($ocrConfig.ContainsKey("SkipMetadataExtract") -and -not $PSBoundParameters.ContainsKey("SkipMetadataExtract")) {
            if ($ocrConfig["SkipMetadataExtract"] -match "^(true|1|yes)$") { $SkipMetadataExtract = $true }
        }
        if ($ocrConfig.ContainsKey("SkipFieldsExtract") -and -not $PSBoundParameters.ContainsKey("SkipFieldsExtract")) {
            if ($ocrConfig["SkipFieldsExtract"] -match "^(true|1|yes)$") { $SkipFieldsExtract = $true }
        }

        # Recalculate langStr if Language was overridden
        if (-not $PSBoundParameters.ContainsKey("Language") -and -not $ocrConfig.ContainsKey("Language")) {
            $langStr = $Language -join "+"
        }
        elseif ($PSBoundParameters.ContainsKey("Language")) {
            $langStr = $Language -join "+"
        }

        # Skip logic based on timestamps
        $shouldSkip = $false
        if (Test-Path -LiteralPath $outPath) {
            $outFileItem = Get-Item -LiteralPath $outPath
            if ($outFileItem.LastWriteTime -ge $file.LastWriteTime) {
                Write-Info "Skipping $($file.Name) - OCR file already exists and is up-to-date."
                $shouldSkip = $true
            }
        }
        elseif (Test-Path -LiteralPath $errPath) {
            $errFileItem = Get-Item -LiteralPath $errPath
            if ($errFileItem.LastWriteTime -ge $file.LastWriteTime) {
                Write-Info "Skipping $($file.Name) - error log already exists and is up-to-date. (Fix error then delete log to retry)"
                $shouldSkip = $true
            }
        }

        if ($shouldSkip) {
            continue
        }

        # Clear old error log if it exists before starting this fresh run
        if (Test-Path -LiteralPath $errPath) {
            Remove-Item -LiteralPath $errPath -Force -ErrorAction SilentlyContinue
        }

        $hasError = $false
        $errorMsg = ""

        if ($ext -match "^\.(pdf)$") {
            # Validate if it already contains text strings using Ghostscript's specialized txtwrite core module
            $hasExistingText = $false
            try {
                $txtCheckArgs = @("-q", "-dNODISPLAY", "-dBATCH", "-dNOPAUSE", "-sDEVICE=txtwrite", "-sOutputFile=-", "`"$($file.FullName)`"")
                $txtOutput = & $ghostscript $txtCheckArgs 2>&1
                $txtBlock = ($txtOutput -join "").Trim()
                if ($txtBlock.Length -gt 5) {
                    $hasExistingText = $true
                }
            }
            catch { }

            if ($hasExistingText) {
                Write-Warn "Pre-check Warning: $($file.Name) appears to already contain a text layer."
                if (-not $ForceOCR) {
                    Write-Info "Skipped $($file.Name). If you wish to re-OCR it to pick up missing textual graphics or schemas, bypass this validation using the -ForceOCR parameter."
                    continue
                }
                else {
                    Write-Info "Bypassing text lock (-ForceOCR supplied). Processing $($file.Name) anyway!"
                }
            }

            $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString())
            New-Item -ItemType Directory -Path $tempDir | Out-Null
            
            $metaTxtPath = Join-Path $tempDir "meta.txt"
            $fieldsTxtPath = Join-Path $tempDir "fields.txt"

            if (-not $SkipMetadataExtract) {
                Write-Info "Extracting metadata from PDF..."
                $dumpMetaArgs = @("`"$($file.FullName)`"", "dump_data_utf8", "output", "`"$metaTxtPath`"")
                Start-Process -FilePath $pdftk -ArgumentList $dumpMetaArgs -Wait -WindowStyle Hidden | Out-Null
            }

            if (-not $SkipFieldsExtract) {
                Write-Info "Extracting fields from PDF..."
                $dumpFieldsArgs = @("`"$($file.FullName)`"", "dump_data_fields_utf8", "output", "`"$fieldsTxtPath`"")
                Start-Process -FilePath $pdftk -ArgumentList $dumpFieldsArgs -Wait -WindowStyle Hidden | Out-Null
            }
            
            try {
                Write-Info ("Extracting pages with Ghostscript at {0} DPI..." -f $Dpi)
                $gsArgs = @(
                    "-dSAFER",
                    "-dBATCH",
                    "-dNOPAUSE",
                    "-r$Dpi",
                    "-sDEVICE=png16m",
                    "-sOutputFile=`"$tempDir\page_%04d.png`""
                )
                if ($Page -gt 0) {
                    $gsArgs += "-dFirstPage=$Page"
                    $gsArgs += "-dLastPage=$Page"
                }
                $gsArgs += "`"$($file.FullName)`""
                if ($GhostscriptArgs) { $gsArgs += $GhostscriptArgs }
                
                $gsProcess = Start-Process -FilePath $ghostscript -ArgumentList $gsArgs -Wait -WindowStyle Hidden -PassThru
                if ($gsProcess.ExitCode -ne 0) {
                    $hasError = $true
                    $errorMsg = "Ghostscript failed to extract images with exit code $($gsProcess.ExitCode)."
                }
                
                if (-not $hasError) {
                    $images = @(Get-ChildItem -Path $tempDir -Filter "*.png" | Sort-Object Name)
                    if ($images.Count -eq 0) {
                        $hasError = $true
                        $errorMsg = "Ghostscript completed but no images were extracted. Ensure Ghostscript handles this PDF properly."
                    }
                    else {
                        if ($PSVersionTable.PSVersion.Major -ge 7) {
                            Write-Info ("OCRing {0} pages in parallel (ThrottleLimit: {1})..." -f $images.Count, $ThrottleLimit)
                            $results = $images | ForEach-Object -Parallel {
                                $img = $_
                                $tesseract = $using:tesseract
                                $langStr = $using:langStr
                                $Dpi = $using:Dpi
                                $TesseractArgs = $using:TesseractArgs

                                $pageRegex = [regex]::Match($img.BaseName, '\d+')
                                $pageNum = if ($pageRegex.Success) { [int]$pageRegex.Value } else { 0 }
                                
                                $outPdfBase = Join-Path $img.DirectoryName $img.BaseName
                                
                                $tessArgs = @(
                                    "`"$($img.FullName)`"",
                                    "`"$outPdfBase`"",
                                    "-l", $langStr,
                                    "--dpi", [string]$Dpi,
                                    "pdf", "txt"
                                )
                                if ($TesseractArgs) { $tessArgs += $TesseractArgs }
                                
                                # Per-page retry logic: up to 2 retries
                                $maxRetries = 2
                                $lastExitCode = -1
                                for ($retry = 0; $retry -le $maxRetries; $retry++) {
                                    if ($retry -gt 0) { Start-Sleep -Seconds 1 }
                                    $tessProcess = Start-Process -FilePath $tesseract -ArgumentList $tessArgs -Wait -WindowStyle Hidden -PassThru
                                    $lastExitCode = $tessProcess.ExitCode
                                    if ($lastExitCode -eq 0) { break }
                                }
                                if ($lastExitCode -ne 0) {
                                    return [PSCustomObject]@{ Success = $false; ErrorMsg = "Tesseract failed on $($img.Name) with exit code $lastExitCode after $($maxRetries + 1) attempts." }
                                }
                                
                                # Append page number header to the text file
                                $txtPath = "$outPdfBase.txt"
                                if (Test-Path -LiteralPath $txtPath) {
                                    $txtContent = Get-Content -LiteralPath $txtPath -Raw
                                    $newContent = "--- PAGE $pageNum ---`r`n" + $txtContent
                                    Set-Content -LiteralPath $txtPath -Value $newContent
                                }
                                
                                return [PSCustomObject]@{ Success = $true; PdfPath = "`"$outPdfBase.pdf`""; TxtPath = $txtPath }
                            } -ThrottleLimit $ThrottleLimit
                        }
                        else {
                            Write-Info ("OCRing {0} pages sequentially (PowerShell 5 - upgrade to PS7 for parallel processing)..." -f $images.Count)
                            $results = foreach ($img in $images) {
                                $pageRegex = [regex]::Match($img.BaseName, '\d+')
                                $pageNum = if ($pageRegex.Success) { [int]$pageRegex.Value } else { 0 }
                                
                                $outPdfBase = Join-Path $img.DirectoryName $img.BaseName
                                
                                $tessArgs = @(
                                    "`"$($img.FullName)`"",
                                    "`"$outPdfBase`"",
                                    "-l", $langStr,
                                    "--dpi", [string]$Dpi,
                                    "pdf", "txt"
                                )
                                if ($TesseractArgs) { $tessArgs += $TesseractArgs }
                                
                                # Per-page retry logic: up to 2 retries
                                $maxRetries = 2
                                $lastExitCode = -1
                                for ($retry = 0; $retry -le $maxRetries; $retry++) {
                                    if ($retry -gt 0) { Start-Sleep -Seconds 1 }
                                    $tessProcess = Start-Process -FilePath $tesseract -ArgumentList $tessArgs -Wait -WindowStyle Hidden -PassThru
                                    $lastExitCode = $tessProcess.ExitCode
                                    if ($lastExitCode -eq 0) { break }
                                }
                                if ($lastExitCode -ne 0) {
                                    [PSCustomObject]@{ Success = $false; ErrorMsg = "Tesseract failed on $($img.Name) with exit code $lastExitCode after $($maxRetries + 1) attempts." }
                                    continue
                                }
                                
                                # Append page number header to the text file
                                $txtPath = "$outPdfBase.txt"
                                if (Test-Path -LiteralPath $txtPath) {
                                    $txtContent = Get-Content -LiteralPath $txtPath -Raw
                                    $newContent = "--- PAGE $pageNum ---`r`n" + $txtContent
                                    Set-Content -LiteralPath $txtPath -Value $newContent
                                }
                                
                                [PSCustomObject]@{ Success = $true; PdfPath = "`"$outPdfBase.pdf`""; TxtPath = $txtPath }
                            }
                        }
                        
                        $failed = $results | Where-Object { -not $_.Success }
                        if ($failed.Count -gt 0) {
                            $hasError = $true
                            $errorMsg = $failed[0].ErrorMsg
                        }
                        
                        if (-not $hasError) {
                            $pdfPages = $results | Select-Object -ExpandProperty PdfPath
                            $txtPages = $results | Select-Object -ExpandProperty TxtPath
                            
                            Write-Info ("Merging {0} OCR'd pages back into PDF..." -f $pdfPages.Count)
                            
                            $allPdftkArgs = @()
                            $allPdftkArgs += $pdfPages
                            $allPdftkArgs += "cat"
                            $allPdftkArgs += "output"
                            
                            # If a watermark is requested, output to a temporary PDF, then watermark it
                            if ($WatermarkPdf -and (Test-Path -LiteralPath $WatermarkPdf)) {
                                $tempMergedPdf = Join-Path $tempDir "temp_merged.pdf"
                                $allPdftkArgs += "`"$tempMergedPdf`""
                                if ($PdftkArgs) { $allPdftkArgs += $PdftkArgs }
                                
                                $pdftkProcess = Start-Process -FilePath $pdftk -ArgumentList $allPdftkArgs -Wait -WindowStyle Hidden -PassThru
                                if ($pdftkProcess.ExitCode -ne 0) {
                                    $hasError = $true
                                    $errorMsg = "PDFtk failed to merge the PDFs with exit code $($pdftkProcess.ExitCode)."
                                }
                                else {
                                    Write-Info "Applying watermark ($WatermarkPdf)..."
                                    $watermarkArgs = @(
                                        "`"$tempMergedPdf`"",
                                        "multibackground",
                                        "`"$WatermarkPdf`"",
                                        "output",
                                        "`"$outPath`""
                                    )
                                    $wmProcess = Start-Process -FilePath $pdftk -ArgumentList $watermarkArgs -Wait -WindowStyle Hidden -PassThru
                                    if ($wmProcess.ExitCode -ne 0) {
                                        $hasError = $true
                                        $errorMsg = "PDFtk failed to apply watermark with exit code $($wmProcess.ExitCode)."
                                    }
                                }
                            }
                            else {
                                $allPdftkArgs += "`"$outPath`""
                                if ($PdftkArgs) { $allPdftkArgs += $PdftkArgs }
                                
                                $pdftkProcess = Start-Process -FilePath $pdftk -ArgumentList $allPdftkArgs -Wait -WindowStyle Hidden -PassThru
                                if ($pdftkProcess.ExitCode -ne 0) {
                                    $hasError = $true
                                    $errorMsg = "PDFtk failed to merge the PDFs with exit code $($pdftkProcess.ExitCode)."
                                }
                            }
                            # Apply metadata to merged output
                            if (-not $hasError -and -not $SkipMetadataExtract -and (Test-Path -LiteralPath $metaTxtPath)) {
                                Write-Info "Applying metadata to output PDF..."
                                $metaPdfPath = Join-Path $tempDir "meta_final.pdf"
                                $metaArgs = @(
                                    "`"$outPath`"",
                                    "update_info_utf8",
                                    "`"$metaTxtPath`"",
                                    "output",
                                    "`"$metaPdfPath`""
                                )
                                $metaProc = Start-Process -FilePath $pdftk -ArgumentList $metaArgs -Wait -WindowStyle Hidden -PassThru
                                if ($metaProc.ExitCode -eq 0 -and (Test-Path -LiteralPath $metaPdfPath)) {
                                    Move-Item -LiteralPath $metaPdfPath -Destination $outPath -Force
                                }
                            }
                            
                            # Merge text files
                            if (-not $hasError) {
                                $outTxtPath = Join-Path $file.DirectoryName ($file.BaseName + "_ocr.txt")
                                Get-Content -LiteralPath $txtPages | Set-Content -LiteralPath $outTxtPath
                                
                                if (-not $SkipMetadataExtract -and (Test-Path -LiteralPath $metaTxtPath)) {
                                    Add-Content -LiteralPath $outTxtPath -Value "`r`n`r`n--- PDF METADATA ---`r`n"
                                    Get-Content -LiteralPath $metaTxtPath | Add-Content -LiteralPath $outTxtPath
                                }
                                
                                if (-not $SkipFieldsExtract -and (Test-Path -LiteralPath $fieldsTxtPath)) {
                                    Add-Content -LiteralPath $outTxtPath -Value "`r`n`r`n--- PDF FORM FIELDS ---`r`n"
                                    Get-Content -LiteralPath $fieldsTxtPath | Add-Content -LiteralPath $outTxtPath
                                }

                                Write-Info "Saved extracted text to $outTxtPath"
                            }
                        }
                    }
                }
            }
            catch {
                $hasError = $true
                $errorMsg = "Unexpected exception occurred: $_"
            }
            finally {
                Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                Write-Info "Cleaned up temporary workspace."
            }
        }
        elseif ($ext -match "^\.(png|jpg|jpeg|tif|tiff|bmp)$") {
            try {
                Write-Info ("OCRing image {0}..." -f $file.Name)
                $outPdfBaseDirect = Join-Path $file.DirectoryName ($file.BaseName + "_ocr")
                
                $tessArgs = @(
                    "`"$($file.FullName)`"",
                    "`"$outPdfBaseDirect`"",
                    "-l", $langStr,
                    "--dpi", [string]$Dpi,
                    "pdf", "txt"
                )
                if ($TesseractArgs) { $tessArgs += $TesseractArgs }
                
                $tessProcess = Start-Process -FilePath $tesseract -ArgumentList $tessArgs -Wait -WindowStyle Hidden -PassThru
                if ($tessProcess.ExitCode -ne 0) {
                    $hasError = $true
                    $errorMsg = "Tesseract failed directly with exit code $($tessProcess.ExitCode)."
                }
                else {
                    if ($WatermarkPdf -and (Test-Path -LiteralPath $WatermarkPdf)) {
                        $tempPdf = "$outPdfBaseDirect.pdf"
                        $tempWMPdf = "$outPdfBaseDirect.temp.pdf"
                        Rename-Item -LiteralPath $tempPdf -NewName (Split-Path $tempWMPdf -Leaf)
                        
                        $watermarkArgs = @(
                            "`"$tempWMPdf`"",
                            "multibackground",
                            "`"$WatermarkPdf`"",
                            "output",
                            "`"$tempPdf`""
                        )
                        $wmProcess = Start-Process -FilePath $pdftk -ArgumentList $watermarkArgs -Wait -WindowStyle Hidden -PassThru
                        if ($wmProcess.ExitCode -ne 0) {
                            $hasError = $true
                            $errorMsg = "PDFtk failed to apply watermark to image with exit code $($wmProcess.ExitCode)."
                        }
                        Remove-Item -LiteralPath $tempWMPdf -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            catch {
                $hasError = $true
                $errorMsg = "Unexpected exception occurred: $_"
            }
        }
        else {
            Write-Warn "Unsupported file extension: $ext for file $($file.FullName)"
            continue
        }

        # Handle the final result for this file
        $numPages = 1
        $numChars = 0
        $numWords = 0
        if ($ext -match "pdf") {
            if ($pdfPages) { $numPages = $pdfPages.Count }
        }

        if ($hasError) {
            Write-Err "Failed to fully process $($file.FullName): $errorMsg"
            Set-Content -LiteralPath $errPath -Value $errorMsg
            Show-ErrorPopup "OCR Error" "Failed to process $($file.FullName): $errorMsg"
            Write-SystemLog "Failed to process $($file.FullName). Error: $errorMsg" "Error"
            
            $reportData += [PSCustomObject]@{
                FullName = $file.FullName
                ReturnCode = 1
                ErrorMessages = $errorMsg
                Pages = $numPages
                Characters = $numChars
                Words = $numWords
            }
        }
        else {
            $elapsed = (Get-Date) - $startTime
            $timeStr = [math]::Round($elapsed.TotalSeconds, 2)
            Write-Info "Success: Created $($file.BaseName)_ocr.pdf in ${timeStr}s"
            Write-SystemLog "Successfully processed $($file.FullName) in ${timeStr} seconds." "Information"
            
            # File Movement Variables
            $outTxtPath = Join-Path $file.DirectoryName ($file.BaseName + "_ocr.txt")
            
            if (Test-Path $outTxtPath) {
                $txtContent = Get-Content $outTxtPath -Raw -ErrorAction SilentlyContinue
                if ($txtContent) {
                    $numChars = $txtContent.Length
                    $numWords = ([regex]::Matches($txtContent, '\w+')).Count
                }
            }

            $reportData += [PSCustomObject]@{
                FullName = $file.FullName
                ReturnCode = 0
                ErrorMessages = ""
                Pages = $numPages
                Characters = $numChars
                Words = $numWords
            }
            
            # Emailing
            if ($EmailTo -and $SmtpServer) {
                Write-Info "Sending email to $($EmailTo -join ', ')..."
                try {
                    $attachments = @()
                    if ($EmailFiles -contains "Source") { $attachments += $file.FullName }
                    if ($EmailFiles -contains "Ocr") { $attachments += $outPath }
                    if ($EmailFiles -contains "Txt" -and (Test-Path $outTxtPath)) { $attachments += $outTxtPath }

                    # Build template variables for email interpolation
                    # Includes file info, processing results, AND all configurable parameters
                    $templateVars = @{
                        # File information
                        filename      = $file.Name
                        basename      = $file.BaseName
                        fullname      = $file.FullName
                        directory     = $file.DirectoryName
                        extension     = $file.Extension
                        # Processing results
                        elapsed       = $timeStr
                        outpath       = $outPath
                        outtxt        = $outTxtPath
                        date          = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                        # All configurable parameters
                        dpi           = [string]$Dpi
                        Language      = $langStr
                        Page          = [string]$Page
                        ThrottleLimit = [string]$ThrottleLimit
                        WatermarkPdf  = if ($WatermarkPdf) { $WatermarkPdf } else { "" }
                        MoveSourceDir = if ($MoveSourceDir) { $MoveSourceDir } else { "" }
                        MoveOcrDir    = if ($MoveOcrDir) { $MoveOcrDir } else { "" }
                        MoveTxtDir    = if ($MoveTxtDir) { $MoveTxtDir } else { "" }
                        RemoveSource  = [string]$RemoveSource
                        ForceOCR      = [string]$ForceOCR
                        EmailTo       = ($EmailTo -join ", ")
                        EmailFiles    = ($EmailFiles -join ", ")
                        SmtpServer    = if ($SmtpServer) { $SmtpServer } else { "" }
                        SmtpPort      = [string]$SmtpPort
                        SmtpUser      = if ($SmtpUser) { $SmtpUser } else { "" }
                    }

                    $mailParams = @{
                        To          = $EmailTo
                        From        = (Expand-Template $EmailFrom $templateVars)
                        Subject     = (Expand-Template $EmailSubject $templateVars)
                        Body        = (Expand-Template $EmailBody $templateVars)
                        SmtpServer  = $SmtpServer
                        Port        = $SmtpPort
                        Attachments = $attachments
                    }
                    if ($EmailReplyTo) {
                        $mailParams.ReplyTo = (Expand-Template $EmailReplyTo $templateVars)
                    }
                    if ($SmtpUser -and $SmtpPassword) {
                        $secPassword = ConvertTo-SecureString $SmtpPassword -AsPlainText -Force
                        $mailParams.Credential = New-Object System.Management.Automation.PSCredential ($SmtpUser, $secPassword)
                    }
                    Send-MailMessage @mailParams -ErrorAction Stop
                }
                catch {
                    Write-Warn "Failed to send email: $_"
                    Write-SystemLog "Failed to send email for $($file.FullName): $_" "Warning"
                }
            }
            
            # File Movement Logic
            if ($MoveOcrDir -and (Test-Path -LiteralPath $MoveOcrDir)) {
                Move-Item -LiteralPath $outPath -Destination $MoveOcrDir -Force
            }
            if ($MoveTxtDir -and (Test-Path -LiteralPath $MoveTxtDir) -and (Test-Path -LiteralPath $outTxtPath)) {
                Move-Item -LiteralPath $outTxtPath -Destination $MoveTxtDir -Force
            }
            if ($MoveSourceDir -and (Test-Path -LiteralPath $MoveSourceDir)) {
                Move-Item -LiteralPath $file.FullName -Destination $MoveSourceDir -Force
            }
            elseif ($RemoveSource) {
                Remove-Item -LiteralPath $file.FullName -Force
            }
        }
        
        # Report Export and Display
        if ($reportData.Count -gt 0) {
            if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
                $reportData | Export-Csv -Path $ReportPath -NoTypeInformation -Force
            }
            if (-not $Quiet -and -not $Silent) {
                Write-Host "`n=== OCR Processing Report ===" -ForegroundColor Cyan
                $reportData | Format-Table -AutoSize | Out-String | Write-Host
            }
        }
    }
}
}

Export-ModuleMember -Function Invoke-OcrTask

