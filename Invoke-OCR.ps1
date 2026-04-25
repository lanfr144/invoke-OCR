<#
.SYNOPSIS
    Performs intelligent OCR on PDF documents or images using Tesseract, Ghostscript, and PDFtk.

.DESCRIPTION
    This script accepts a list of file paths (images or PDFs), either via parameter or pipeline.
    It intelligently locates required software (Ghostscript, Tesseract, PDFtk) automatically.
    
    If the file is an image, it runs Tesseract directly to produce a searchable PDF.
    If the file is a PDF, it checks if text already exists. If text exists, it warns the user 
    and skips it, unless the -ForceOCR switch is passed. It explodes the PDF into high resolution 
    images using Ghostscript, OCRs each page using Tesseract, and merges them back together 
    into a single PDF using PDFtk. 

    The resulting file will end with "_ocr.pdf". It can also automatically move files to
    archive directories and email the results via SMTP.

    PARALLEL PROCESSING
    On PowerShell 7+, pages are OCR'd in parallel using ForEach-Object -Parallel.
    On PowerShell 5.1, pages are processed sequentially (consider upgrading for speed).

    PER-DIRECTORY CONFIGURATION (.ocrconfig)
    You can place a ".ocrconfig" file in the same directory as your source files to override
    default parameters without passing them on the command line. The config file uses a simple
    Key=Value format. Command-line parameters always take precedence over config file values.

    Supported .ocrconfig keys:
        Language      = fra+eng
        Dpi           = 600
        ForceOCR      = true
        WatermarkPdf  = C:\watermarks\header.pdf
        MoveSourceDir = C:\Archive\Originals
        MoveOcrDir    = C:\Archive\Processed
        MoveTxtDir    = C:\Archive\Text
        RemoveSource  = true
        EmailTo       = user@example.com
        EmailSubject  = OCR Done: ${filename}
        EmailBody     = File "${filename}" processed at ${dpi} DPI with ${Language}. Took ${elapsed}s.
        EmailFrom     = Scanner <scanner@company.com>
        EmailReplyTo  = admin@company.com
        SmtpServer    = smtp.company.local
        SmtpPort      = 587
        SmtpUser      = scanner
        SmtpPassword  = secret

    Lines starting with # or ' are treated as comments. Values can be quoted or unquoted.

    EMAIL TEMPLATE VARIABLES
    The parameters EmailSubject, EmailBody, EmailFrom, and EmailReplyTo support template
    variables that are expanded at runtime. Available placeholders:

        ${filename}      - Source file name (e.g. invoice.pdf)
        ${basename}      - File name without extension (e.g. invoice)
        ${fullname}      - Full absolute path of the source file
        ${directory}     - Directory containing the source file
        ${extension}     - File extension including dot (e.g. .pdf)
        ${elapsed}       - Processing time in seconds
        ${outpath}       - Path to the generated _ocr.pdf
        ${outtxt}        - Path to the generated _ocr.txt
        ${date}          - Current date/time in yyyy-MM-dd HH:mm:ss format
        ${dpi}           - DPI value used for processing
        ${Language}      - Language string (e.g. eng+fra+deu)
        ${Page}          - Page number processed (0 = all)
        ${ThrottleLimit} - Parallel thread count
        ${WatermarkPdf}  - Watermark PDF path (if set)
        ${MoveSourceDir} - Source archive directory (if set)
        ${MoveOcrDir}    - OCR output archive directory (if set)
        ${MoveTxtDir}    - Text output archive directory (if set)
        ${RemoveSource}  - Whether source is deleted (True/False)
        ${ForceOCR}      - Whether ForceOCR was active (True/False)
        ${EmailTo}       - Recipient addresses
        ${EmailFiles}    - Attached file types
        ${SmtpServer}    - SMTP server used
        ${SmtpPort}      - SMTP port used
        ${SmtpUser}      - SMTP username (if set)

.PARAMETER Path
    The file path or array of file paths to process. Accepts pipeline input.
    You can also pipe FileInfo objects from Get-ChildItem.

.PARAMETER Language
    One or many Tesseract language codes (e.g., eng, fra, deu). Default is eng+fra+deu+ltz+por+lat.
    See available languages: tesseract --list-langs

.PARAMETER Dpi
    DPI for PDF bursting via Ghostscript and Tesseract OCR. Default is 300.
    Higher values improve quality but increase processing time and file size.

.PARAMETER Page
    Specific page to process. If 0 or omitted, processes all pages.

.PARAMETER WatermarkPdf
    Path to a PDF file to use as a background watermark (philigram).
    Applied via PDFtk's multibackground command after OCR.

.PARAMETER ThrottleLimit
    Number of pages to process concurrently in Tesseract. Requires PowerShell 7. Default is 4.
    Set higher on machines with many CPU cores for faster processing.

.PARAMETER Yes
    Bypasses the confirmation prompt asking if the script found the correct executables.
    Alias: -y

.PARAMETER Quiet
    Suppresses standard output text (like "Extracting pages..."), but still prints errors/warnings.

.PARAMETER Silent
    Suppresses EVERYTHING. No success messages, no warnings, no errors to the console.
    Designed for background/scheduled task usage.

.PARAMETER ForceOCR
    If the PDF already has a text layer, process it anyway. Useful for documents with
    schemas or images that still need OCR despite having some existing text.

.PARAMETER TesseractPath
    Explicit path to the Tesseract executable. Overrides automatic discovery.

.PARAMETER GhostscriptPath
    Explicit path to the Ghostscript executable. Overrides automatic discovery.

.PARAMETER PdftkPath
    Explicit path to the PDFtk executable. Overrides automatic discovery.

.PARAMETER TesseractArgs
    Additional custom arguments to pass verbatim to Tesseract.

.PARAMETER GhostscriptArgs
    Additional custom arguments to pass verbatim to Ghostscript.

.PARAMETER PdftkArgs
    Additional custom arguments to pass verbatim to PDFtk.

.PARAMETER MoveSourceDir
    Directory to move the original PDF/image to after a successful scan.

.PARAMETER MoveOcrDir
    Directory to move the generated _ocr.pdf to after a successful scan.

.PARAMETER MoveTxtDir
    Directory to move the generated _ocr.txt to after a successful scan.

.PARAMETER RemoveSource
    Permanently deletes the original PDF/image after a successful scan.

.PARAMETER EmailTo
    Array of email addresses to send results to (e.g., "Mr Smith <a@b.com>").

.PARAMETER EmailFiles
    Which files to attach to the email. Options: Source, Ocr, Txt. Default is Ocr.

.PARAMETER EmailSubject
    Email subject line. Supports template variables (see DESCRIPTION).
    Default: "OCR Completed: ${filename}"

.PARAMETER EmailBody
    Email body text. Supports template variables (see DESCRIPTION).
    Default: "The document ${filename} was successfully processed in ${elapsed} seconds."

.PARAMETER EmailFrom
    Sender address for the email. Supports template variables.
    Default: "Invoke-OCR <no-reply@localhost>"

.PARAMETER EmailReplyTo
    Reply-To address for the email. Supports template variables.

.PARAMETER SmtpServer
    SMTP Server address required to send emails.

.PARAMETER SmtpPort
    Port for the SMTP server. Default is 25.

.PARAMETER SmtpUser
    Username for SMTP authentication.

.PARAMETER SmtpPassword
    Password for SMTP authentication.

.EXAMPLE
    .\Invoke-OCR.ps1 -Path "document.pdf"

    Basic usage: OCR a single PDF with default settings (6 languages, 300 DPI).

.EXAMPLE
    Get-ChildItem -Filter "*.pdf" | .\Invoke-OCR.ps1 -Language "eng","fra" -y -Silent -ForceOCR

    Automation: Process all PDFs in current folder silently with English+French only.

.EXAMPLE
    .\Invoke-OCR.ps1 -Path "invoice.pdf" -MoveSourceDir "C:\Archive\Originals" -MoveOcrDir "C:\Archive\Processed" -EmailTo "finance@company.com" -SmtpServer "smtp.company.local"

    Process, archive, and email the results.

.EXAMPLE
    .\Invoke-OCR.ps1 -Path "report.pdf" -EmailTo "boss@company.com" -EmailSubject "Scan ready: ${filename}" -EmailBody "Hi, the file ${filename} (${basename}) was scanned at ${dpi} DPI using ${Language}. Processing took ${elapsed}s." -SmtpServer "smtp.local"

    Custom email template with variable interpolation.

.EXAMPLE
    .\Invoke-OCR.ps1 -Path "scan.pdf" -Dpi 600 -Language "deu" -ThrottleLimit 8 -ForceOCR

    High-quality German-only OCR at 600 DPI with 8 parallel threads, forcing re-OCR.

.NOTES
    Author  : Invoke-OCR Project
    Requires: PowerShell 5.1+ (PowerShell 7+ recommended for parallel processing)

    PREREQUISITES - The following tools must be installed:

    1. Tesseract OCR - The core OCR engine
       Download : https://github.com/UB-Mannheim/tesseract/wiki
       Man page : https://tesseract-ocr.github.io/tessdoc/Command-Line-Usage.html
       Reference: https://github.com/tesseract-ocr/tesseract/blob/main/doc/tesseract.1.asc

    2. Ghostscript - PDF rendering and text detection
       Download : https://ghostscript.com/releases/gsdnld.html
       Docs     : https://ghostscript.com/docs/9.54.0/Use.htm

    3. PDFtk Server - PDF merging and watermarking
       Download : https://www.pdflabs.com/tools/pdftk-server/
       Docs     : https://www.pdflabs.com/docs/pdftk-man-page/

    4. PowerShell 7+ (optional, for parallel processing)
       Download : https://aka.ms/powershell

.LINK
    https://github.com/UB-Mannheim/tesseract/wiki

.LINK
    https://ghostscript.com/releases/gsdnld.html

.LINK
    https://www.pdflabs.com/tools/pdftk-server/

.LINK
    https://tesseract-ocr.github.io/tessdoc/Command-Line-Usage.html

.LINK
    https://www.pdflabs.com/docs/pdftk-man-page/
#>

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
    [string]$SmtpPassword
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
        if (-not $Silent) { Write-Error -Message $Message }
    }

    function Expand-Template {
        param([string]$Template, [hashtable]$Variables)
        $result = $Template
        foreach ($key in $Variables.Keys) {
            $result = $result -replace [regex]::Escape("`${$key}"), $Variables[$key]
        }
        return $result
    }

    function Import-OcrConfig {
        param([string]$ConfigPath)
        $config = @{}
        if (-not (Test-Path -LiteralPath $ConfigPath)) { return $config }

        $lines = Get-Content -LiteralPath $ConfigPath
        foreach ($line in $lines) {
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
                $config[$key] = $value
            }
        }
        Write-Info "Loaded .ocrconfig from $ConfigPath ($($config.Count) settings)"
        return $config
    }

    function Write-SystemLog {
        param([string]$Message, [string]$Type = "Information")
        $logPath = Join-Path $PSScriptRoot "ocr_service.log"
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

    function Show-ErrorPopup {
        param([string]$Title, [string]$Message)
        try {
            Start-Process "msg.exe" -ArgumentList "* `"${Title}: $Message`"" -WindowStyle Hidden
        }
        catch { }
    }

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
                        # Process each page with Tesseract (parallel on PS7+, sequential on PS5)
                        $ocrScriptBlock = {
                            param($img, $tesseract, $langStr, $Dpi, $TesseractArgs)
                            
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
                            
                            $tessProcess = Start-Process -FilePath $tesseract -ArgumentList $tessArgs -Wait -NoNewWindow -PassThru
                            if ($tessProcess.ExitCode -ne 0) {
                                return [PSCustomObject]@{ Success = $false; ErrorMsg = "Tesseract failed on $($img.Name) with exit code $($tessProcess.ExitCode)." }
                            }
                            
                            # Append page number header to the text file
                            $txtPath = "$outPdfBase.txt"
                            if (Test-Path -LiteralPath $txtPath) {
                                $txtContent = Get-Content -LiteralPath $txtPath -Raw
                                $newContent = "--- PAGE $pageNum ---`r`n" + $txtContent
                                Set-Content -LiteralPath $txtPath -Value $newContent
                            }
                            
                            return [PSCustomObject]@{ Success = $true; PdfPath = "`"$outPdfBase.pdf`""; TxtPath = $txtPath }
                        }

                        if ($PSVersionTable.PSVersion.Major -ge 7) {
                            Write-Info ("OCRing {0} pages in parallel (ThrottleLimit: {1})..." -f $images.Count, $ThrottleLimit)
                            $results = $images | ForEach-Object -Parallel {
                                $ocrBlock = $using:ocrScriptBlock
                                & $ocrBlock $_ $using:tesseract $using:langStr $using:Dpi $using:TesseractArgs
                            } -ThrottleLimit $ThrottleLimit
                        }
                        else {
                            Write-Info ("OCRing {0} pages sequentially (PowerShell 5 - upgrade to PS7 for parallel processing)..." -f $images.Count)
                            $results = foreach ($img in $images) {
                                & $ocrScriptBlock $img $tesseract $langStr $Dpi $TesseractArgs
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
                                
                                $pdftkProcess = Start-Process -FilePath $pdftk -ArgumentList $allPdftkArgs -Wait -NoNewWindow -PassThru
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
                                    $wmProcess = Start-Process -FilePath $pdftk -ArgumentList $watermarkArgs -Wait -NoNewWindow -PassThru
                                    if ($wmProcess.ExitCode -ne 0) {
                                        $hasError = $true
                                        $errorMsg = "PDFtk failed to apply watermark with exit code $($wmProcess.ExitCode)."
                                    }
                                }
                            }
                            else {
                                $allPdftkArgs += "`"$outPath`""
                                if ($PdftkArgs) { $allPdftkArgs += $PdftkArgs }
                                
                                $pdftkProcess = Start-Process -FilePath $pdftk -ArgumentList $allPdftkArgs -Wait -NoNewWindow -PassThru
                                if ($pdftkProcess.ExitCode -ne 0) {
                                    $hasError = $true
                                    $errorMsg = "PDFtk failed to merge the PDFs with exit code $($pdftkProcess.ExitCode)."
                                }
                            }
                            
                            # Merge text files
                            if (-not $hasError) {
                                $outTxtPath = Join-Path $file.DirectoryName ($file.BaseName + "_ocr.txt")
                                Get-Content -LiteralPath $txtPages | Set-Content -LiteralPath $outTxtPath
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
                
                $tessProcess = Start-Process -FilePath $tesseract -ArgumentList $tessArgs -Wait -NoNewWindow -PassThru
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
                        $wmProcess = Start-Process -FilePath $pdftk -ArgumentList $watermarkArgs -Wait -NoNewWindow -PassThru
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
            Write-Warn "Unsupported file extension: $ext for file $($file.Name)"
            continue
        }

        # Handle the final result for this file
        if ($hasError) {
            Write-Err "Failed to fully process $($file.Name): $errorMsg"
            Set-Content -LiteralPath $errPath -Value $errorMsg
            Show-ErrorPopup "OCR Error" "Failed to process $($file.Name): $errorMsg"
            Write-SystemLog "Failed to process $($file.Name). Error: $errorMsg" "Error"
        }
        else {
            $elapsed = (Get-Date) - $startTime
            $timeStr = [math]::Round($elapsed.TotalSeconds, 2)
            Write-Info "Success: Created $($file.BaseName)_ocr.pdf in ${timeStr}s"
            Write-SystemLog "Successfully processed $($file.Name) in ${timeStr} seconds." "Information"
            
            # File Movement Variables
            $outTxtPath = Join-Path $file.DirectoryName ($file.BaseName + "_ocr.txt")
            
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
                    Write-SystemLog "Failed to send email for $($file.Name): $_" "Warning"
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
    }
}
