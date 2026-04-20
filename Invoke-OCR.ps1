<#
.SYNOPSIS
    Performs OCR on PDF documents or images using Tesseract, Ghostscript, and PDFtk.

.DESCRIPTION
    This script accepts a list of file paths (images or PDFs), either via parameter or pipeline.
    It intelligently locates required software (Ghostscript, Tesseract, PDFtk) automatically.
    
    If the file is an image, it runs Tesseract directly to produce a searchable PDF.
    If the file is a PDF, it checks if text already exists. If text exists, it warns the user 
    and skips it, unless the -ForceOCR switch is passed. It explodes the PDF into high resolution 
    images using Ghostscript, OCRs each page using Tesseract, and merges them back together 
    into a single PDF using PDFtk. The resulting file will end with "_ocr.pdf".

.EXAMPLE
    # Basic usage:
    .\Invoke-OCR.ps1 -Path "document.pdf"
    
.EXAMPLE
    # Automation usage avoiding questions and warnings:
    Get-ChildItem -Filter "*.pdf" | .\Invoke-OCR.ps1 -Language "eng+fra" -y -Silent -ForceOCR

.EXAMPLE
    # Dealing with spaces, accents, and hyphens at the beginning of file names:
    .\Invoke-OCR.ps1 -Path 'C:\My Scans\-éxâmple file.pdf'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [Alias("FullName")]
    [string[]]$Path,

    # Supports one or many Tesseract language codes, array format.
    [string[]]$Language = @("eng", "fra", "deu", "ltz", "por", "lat"),

    # DPI for PDF bursting via Ghostscript and Tesseract OCR
    [int]$Dpi = 300,

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
    [string[]]$PdftkArgs
)

BEGIN {
    # Logging flow control wrappers
    function Write-Info {
        param([string]$Message)
        if (-not $Quiet -and -not $Silent) { Write-Host $Message }
    }
    
    function Write-Warn {
        param([string]$Message)
        if (-not $Silent) { Write-Warning $Message }
    }
    
    function Write-Err {
        param([string]$Message)
        if (-not $Silent) { Write-Error $Message }
    }

    function Verify-Executable {
        param([string]$ExePath, [string]$TestArg, [string]$RegexMatch)
        try {
            $output = & $ExePath $TestArg 2>&1
            $outStr = $output -join "`n"
            if ($outStr -match $RegexMatch) { return $true }
        } catch { }
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
    $ghostscript = Find-Executable -ExplicitPath $GhostscriptPath -ExePattern "gswin*.exe" -FolderPattern "gs*" -VerifyArg "-v" -VerifyRegex "(?i)Ghostscript"
    $pdftk = Find-Executable -ExplicitPath $PdftkPath -ExePattern "pdftk*.exe" -FolderPattern "pdftk*" -VerifyArg "--help" -VerifyRegex "(?i)pdftk"

    $halt = $false
    if (-not $tesseract) { Write-Err "Tesseract could not be found via args, PATH, or default directories."; $halt = $true }
    if (-not $ghostscript) { Write-Err "Ghostscript could not be found via args, PATH, or default directories."; $halt = $true }
    if (-not $pdftk) { Write-Err "PDFtk could not be found via args, PATH, or default directories."; $halt = $true }
    
    if ($halt) { throw "Missing required prerequisites." }
    
    Write-Info "Ready. All executables verified."
}

PROCESS {
    $langStr = $Language -join "+"

    foreach ($p in $Path) {
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
            } catch { }

            if ($hasExistingText) {
                Write-Warn "Pre-check Warning: $($file.Name) appears to already contain a text layer."
                if (-not $ForceOCR) {
                    Write-Info "Skipped $($file.Name). If you wish to re-OCR it to pick up missing textual graphics or schemas, bypass this validation using the -ForceOCR parameter."
                    continue
                } else {
                    Write-Info "Bypassing text lock (-ForceOCR supplied). Processing $($file.Name) anyway!"
                }
            }

            $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString())
            New-Item -ItemType Directory -Path $tempDir | Out-Null
            
            try {
                Write-Info ("Extracting {0} pages with Ghostscript at {1} DPI..." -f $file.Name, $Dpi)
                $gsArgs = @(
                    "-dSAFER",
                    "-dBATCH",
                    "-dNOPAUSE",
                    "-r$Dpi",
                    "-sDEVICE=png16m",
                    "-sOutputFile=`"$tempDir\page_%04d.png`"",
                    "`"$($file.FullName)`""
                )
                if ($GhostscriptArgs) { $gsArgs += $GhostscriptArgs }
                
                $gsProcess = Start-Process -FilePath $ghostscript -ArgumentList $gsArgs -Wait -NoNewWindow -PassThru
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
                        $pdfPages = @()
                        $i = 1
                        foreach ($img in $images) {
                            Write-Info ("OCRing page {0} of {1}..." -f $i, $images.Count)
                            $outPdfBase = Join-Path $tempDir $img.BaseName
                            
                            $tessArgs = @(
                                "`"$($img.FullName)`"",
                                "`"$outPdfBase`"",
                                "-l", $langStr,
                                "--dpi", [string]$Dpi,
                                "pdf"
                            )
                            if ($TesseractArgs) { $tessArgs += $TesseractArgs }
                            
                            $tessProcess = Start-Process -FilePath $tesseract -ArgumentList $tessArgs -Wait -NoNewWindow -PassThru
                            if ($tessProcess.ExitCode -ne 0) {
                                $hasError = $true
                                $errorMsg = "Tesseract failed on page $i with exit code $($tessProcess.ExitCode)."
                                break
                            }
                            
                            $pdfPages += "`"$outPdfBase.pdf`""
                            $i++
                        }
                        
                        if (-not $hasError) {
                            Write-Info ("Merging {0} OCR'd pages back into PDF..." -f $pdfPages.Count)
                            $allPdftkArgs = @()
                            $allPdftkArgs += $pdfPages
                            $allPdftkArgs += "cat"
                            $allPdftkArgs += "output"
                            $allPdftkArgs += "`"$outPath`""
                            if ($PdftkArgs) { $allPdftkArgs += $PdftkArgs }
                            
                            $pdftkProcess = Start-Process -FilePath $pdftk -ArgumentList $allPdftkArgs -Wait -NoNewWindow -PassThru
                            if ($pdftkProcess.ExitCode -ne 0) {
                                $hasError = $true
                                $errorMsg = "PDFtk failed to merge the PDFs with exit code $($pdftkProcess.ExitCode)."
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
                    "pdf"
                )
                if ($TesseractArgs) { $tessArgs += $TesseractArgs }
                
                $tessProcess = Start-Process -FilePath $tesseract -ArgumentList $tessArgs -Wait -NoNewWindow -PassThru
                if ($tessProcess.ExitCode -ne 0) {
                    $hasError = $true
                    $errorMsg = "Tesseract failed directly with exit code $($tessProcess.ExitCode)."
                }
            }
            catch {
                $hasError = $true
                $errorMsg = "Unexpected exception occurred: $_"
            }
        }
        else {
            Write-Warn "Unsupported file extension: $ext for file $($file.Name)"
            continue
        }

        # Handle the final result for this file
        if ($hasError) {
            Write-Err "Failed to fully process $($file.Name): $errorMsg"
            Set-Content -LiteralPath $errPath -Value $errorMsg
        }
        else {
            Write-Info "Success: Created $($file.BaseName)_ocr.pdf"
        }
    }
}
