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
    Password for SMTP authentication. Can also be set to "credential:<name>" to use
    a securely stored credential (see Save-OcrCredential.ps1).

.PARAMETER ValidateConfig
    When set, the script validates .ocrconfig files in the target path directories
    and exits without processing any files. Reports valid keys, unknown keys, and
    parent config inheritance. Useful for testing config files before deployment.

.EXAMPLE
    .\Invoke-OCR.ps1 -Path "document.pdf"

    Basic usage: OCR a single PDF with default settings (6 languages, 300 DPI).

.EXAMPLE
    Get-ChildItem -Filter "*.pdf" | .\Invoke-OCR.ps1 -Language "eng","fra" -y -Silent -ForceOCR

    Automation: Process all PDFs in current folder silently with English+French only.

.EXAMPLE
    .\Invoke-OCR.ps1 -Path "C:\scans" -ValidateConfig

    Validate the .ocrconfig file in C:\scans without processing any files.

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
    [string]$SmtpPassword,

    # Validation mode: only validate .ocrconfig files, do not process
    [switch]$ValidateConfig,

    # Path to save the post-process CSV report
    [string]$ReportPath,

    # Skip extracting and injecting PDF metadata
    [switch]$SkipMetadataExtract,

    # Skip extracting PDF form fields
    [switch]$SkipFieldsExtract,

    [switch]$PassThru
)

$moduleRoot = Join-Path $PSScriptRoot 'Modules'
Import-Module (Join-Path $moduleRoot 'OcrLogging.psm1') -Force
Import-Module (Join-Path $moduleRoot 'OcrNotification.psm1') -Force
Import-Module (Join-Path $moduleRoot 'OcrConfig.psm1') -Force
Import-Module (Join-Path $moduleRoot 'OcrEmail.psm1') -Force
Import-Module (Join-Path $moduleRoot 'InvokeOcr.psm1') -Force

Set-OcrLoggingState -Quiet $Quiet -Silent $Silent

$exitCode = 0
try {
    $argsForTask = @{}
    foreach ($k in $PSBoundParameters.Keys) { $argsForTask[$k] = $PSBoundParameters[$k] }
    $argsForTask.Remove("PassThru") | Out-Null

    # Call the core function
    Invoke-OcrTask @argsForTask
}
catch {
    Write-Host "OCR Task Failed: $_" -ForegroundColor Red
    $exitCode = 1
}

if ($PassThru) {
    [PSCustomObject]@{
        ExitCode = $exitCode
        Message  = if ($exitCode -eq 0) { "Success" } else { "Failed" }
    }
}

return $exitCode

