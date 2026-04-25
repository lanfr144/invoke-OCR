# Invoke-OCR 🔍

A robust, intelligent PowerShell automation script designed to seamlessly extract, analyze, and convert static documents (PDFs and images) into fully searchable OCR'd PDFs.

## 🌟 Key Features

- **Intelligent Pre-flight Checks**: Uses `Ghostscript` to invisibly scan PDFs to see if they already contain a text-layer. If they do, it saves you processing time by aggressively skipping them (unless bypassed via `-ForceOCR`).
- **Parallel Processing Engine**: Natively uses PowerShell 7's `ForEach-Object -Parallel` to OCR multiple pages simultaneously, drastically reducing processing time.
- **Per-Page Retry Logic**: If Tesseract fails on a page (corruption, memory spike), the script automatically retries up to 2 times before marking it as failed.
- **Dynamic Path Discovery**: No need to hardcode paths. The script natively hunts for your executables checking your user arguments, scanning your system `$env:PATH`, and recursively digging through your `C:\Program Files` looking for your missing binaries (automatically ignoring uninstallers).
- **Text & Page Generation**: Automatically extracts and compiles a `_ocr.txt` file alongside your PDF, intelligently injecting `--- PAGE X ---` headers to split up the text data.
- **Watermarking (Philigram)**: Built-in support to seamlessly stamp a custom watermark across all your generated pages using PDFtk.
- **Timestamp Logic**: Avoids re-processing the same files over and over. If an `_ocr.pdf` already exists and is newer than the source, the script gracefully moves onto the next task.
- **Fail-safe Error Logging**: A failure doesn't crash the loop. If a specific page corrupts during Ghostscript bursting or Tesseract parsing, the script catches it, isolates a `document.err.log` file, and keeps working through your pipeline.
- **Log Rotation**: The `ocr_service.log` file is automatically rotated when it exceeds 5MB, preventing disk space issues on long-running systems.
- **Automated Emailing**: Natively hooks into an SMTP server with customizable email templates using `${variable}` placeholders. Supports secure credential storage via Windows DPAPI.
- **Smart Archiving**: Move the original source files and output files to separate backup directories automatically, or completely delete the original source when finished to keep your drop folders clean.
- **Windows Event Logging**: Every file processed is securely logged to the native Windows Event Viewer (Application log), recording exact processing times and success/failure states.
- **Desktop Notifications**: Uses native Windows Toast Notifications (BurntToast or .NET API) with automatic fallback to `msg.exe` on older systems.
- **Hierarchical Config Files**: Place `.ocrconfig` files in any directory for per-folder settings. Child directories inherit parent configs, allowing base defaults with local overrides.
- **Config Validation**: Built-in `-ValidateConfig` flag and standalone `Test-OcrConfig.ps1` script to catch typos and invalid keys before deployment.
- **Bulletproof Parsing**: Fully compatible with `-LiteralPath` so filenames with spaces, hyphens (`-`), brackets (`[]`), or foreign accents never break your shell!

## ⚙️ Prerequisites

This script acts as the orchestration layer between three powerful open-source utilities. You must have these installed on your Windows machine:

1. **[Ghostscript](https://ghostscript.com/releases/gsdnld.html)**: Used for exploding static PDFs into high-resolution images, and detecting existing text strings.
2. **[Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)**: The core Optical Character Recognition engine.
3. **[PDFtk Server](https://www.pdflabs.com/tools/pdftk-server/)**: Used to rapidly safely merge the resulting standalone OCR'd pages back into a cohesive, single PDF.

*Note: Parallel processing and background trigger features require **PowerShell 7+**.*

---

## 🤖 Automated Background Trigger System

You can transform this script into a fully automated 24/7 background service! When installed, you can simply drop a file into a specific folder, and it will be silently processed in the background.

The default watch folder is `C:\scans`, but you can customize it with the `-WatchFolder` parameter.
Language-specific subdirectories are created automatically:
- `<WatchFolder>\en` ➔ Maps to `-Language eng`
- `<WatchFolder>\fr` ➔ Maps to `-Language fra`
- `<WatchFolder>\de` ➔ Maps to `-Language deu`
- `<WatchFolder>\lb` ➔ Maps to `-Language ltz`
- `<WatchFolder>` ➔ Base default mapping (all languages)

The watcher features automatic crash recovery (5 retries with exponential backoff) and duplicate event suppression to prevent double-processing.

### Installation & Management

The project includes management scripts to control your background service:

1. **Install the Service**: Run `.\Install-OCRWatcher.ps1 [-WatchFolder "D:\custom"]`
   - *Checks prerequisites, creates folder structure, registers a hidden Scheduled Task.*
2. **Check Status**: Run `.\Get-OCRWatcherStatus.ps1`
   - *Verifies if the listener is currently `RUNNING`, `DISABLED`, or `READY`.*
3. **Uninstall the Service**: Run `.\Remove-OCRWatcher.ps1`
   - *Completely unregisters and deletes the background task from your system.*

---

## 🖱️ Windows Right-Click Integration

You don't have to use the background folders! You can inject the OCR tool directly into Windows Explorer.

1. **Install Menu**: Run `.\Install-ContextMenu.ps1` (it will auto-request Admin rights).
2. **Usage**: You can now right-click ANY PDF, PNG, JPG, or TIFF anywhere on your computer (even your Desktop) and click **"Make Searchable (Invoke-OCR)"**.
3. **Remove Menu**: Run `.\Remove-ContextMenu.ps1`.

---

## 🚀 Manual Usage

If you prefer to run it manually, the script is built natively for PowerShell Advanced Functions, meaning it seamlessly accepts pipeline inputs and parameters.

### Basic Run
Run standard OCR on a single document defaults (6 languages, 300 DPI, 4 parallel threads):
```powershell
.\Invoke-OCR.ps1 -Path "C:\My Scans\document.pdf"
```

### Automation Pipeline
Process an entire folder of PNGs quietly, automatically answering "YES" to the executable prompts:
```powershell
Get-ChildItem -Filter "*.png" | .\Invoke-OCR.ps1 -Quiet -y
```

### Watermarks & Performance Tuning
Apply a custom watermark PDF to a specific file, processing 8 pages at the same time:
```powershell
.\Invoke-OCR.ps1 -Path "confidential.pdf" -WatermarkPdf "C:\watermarks\top_secret.pdf" -ThrottleLimit 8
```

### Force & Advanced Overrides
Process a PDF that *already has some text* (like a schema missing OCR), limit the OCR engine purely to English, and inject raw Tesseract arguments to whitelist numbers:
```powershell
.\Invoke-OCR.ps1 -Path "invoice.pdf" -Language "eng" -ForceOCR -TesseractArgs @("--oem", "1", "-c", "tessedit_char_whitelist=0123456789")
```

## 🎛️ Parameters

| Parameter | Type | Default | Description |
| :--- | :---: | :--- | :--- |
| `-Path` | `String[]` | Required | Target file strings (supports arrays and pipelines) |
| `-Language` | `String[]` | `eng`, `fra`, `deu`, `ltz`, `por`, `lat` | Tesseract language training packages. Array or '+' separated string. |
| `-Dpi` | `Int` | `300` | Target rendering resolution passed to both Ghostscript and Tesseract |
| `-Page` | `Int` | `0` (All) | Process only a specific page number. Overrides full-document bursting. |
| `-WatermarkPdf` | `String` | None | Absolute path to a PDF file. Will be stamped as a background via PDFtk. |
| `-ThrottleLimit`| `Int` | `4` | Number of concurrent pages to process in parallel via Tesseract. |
| `-y` | `Switch` | `$false` | Unattended bypass. Disables the interactive user-prompts confirming executables. |
| `-Quiet` | `Switch` | `$false` | Console verbosity lock. Suppresses success/progress messages but allows Warning & Errors. |
| `-Silent` | `Switch` | `$false` | The nuclear option. Mutes everything. Designed for background scheduled tasks. |
| `-ForceOCR` | `Switch` | `$false` | Unlocks files that fail the `txtwrite` pre-flight check, processing them regardless of existing text. |
| `-MoveSourceDir`| `String` | None | Directory to move the original PDF/image to after a successful scan. |
| `-MoveOcrDir`   | `String` | None | Directory to move the generated `_ocr.pdf` to. |
| `-MoveTxtDir`   | `String` | None | Directory to move the generated `_ocr.txt` to. |
| `-RemoveSource` | `Switch` | `$false` | Permanently deletes the original PDF/image after a successful scan. |
| `-EmailTo`      | `String[]`| None | Array of email addresses (e.g., `"Mr Smith" <a@b.com>`) to send results to. |
| `-EmailFiles`   | `String[]`| `Ocr` | Which files to attach. Options: `Source`, `Ocr`, `Txt`. |
| `-EmailSubject` | `String` | `OCR Completed: ${filename}` | Email subject line. Supports template variables (see below). |
| `-EmailBody`    | `String` | `The document ${filename} was successfully processed in ${elapsed} seconds.` | Email body. Supports template variables. |
| `-EmailFrom`    | `String` | `Invoke-OCR <no-reply@localhost>` | Sender address. Supports template variables. |
| `-EmailReplyTo` | `String` | None | Reply-To address. Supports template variables. |
| `-SmtpServer`   | `String` | None | SMTP Server address required to send emails. |
| `-SmtpPort`     | `Int`    | `25` | Port for the SMTP server. |
| `-SmtpUser`     | `String` | None | Username for SMTP authentication. |
| `-SmtpPassword` | `String` | None | Password for SMTP auth. Use `credential:<name>` for secure DPAPI storage. |
| `-ValidateConfig` | `Switch` | `$false` | Validate `.ocrconfig` files only — do not process any files. |

*You can also directly inject overriding paths via `-TesseractPath`, `-GhostscriptPath`, and `-PdftkPath`.*

---

## 📧 Email Template Variables

The `-EmailSubject`, `-EmailBody`, `-EmailFrom`, and `-EmailReplyTo` parameters support dynamic placeholders that are expanded at runtime:

| Variable | Expands To | Example |
| :--- | :--- | :--- |
| `${filename}` | Source file name | `invoice.pdf` |
| `${basename}` | File name without extension | `invoice` |
| `${fullname}` | Full absolute path | `C:\scans\invoice.pdf` |
| `${directory}` | Directory containing the file | `C:\scans` |
| `${extension}` | File extension (with dot) | `.pdf` |
| `${elapsed}` | Processing time in seconds | `12.34` |
| `${outpath}` | Path to generated `_ocr.pdf` | `C:\scans\invoice_ocr.pdf` |
| `${outtxt}` | Path to generated `_ocr.txt` | `C:\scans\invoice_ocr.txt` |
| `${date}` | Current date/time | `2026-04-25 08:30:00` |
| `${dpi}` | DPI value used | `300` |
| `${Language}` | Language string | `eng+fra+deu` |
| `${Page}` | Page number (0 = all) | `0` |
| `${ThrottleLimit}` | Parallel thread count | `4` |
| `${WatermarkPdf}` | Watermark PDF path | `C:\watermarks\header.pdf` |
| `${MoveSourceDir}` | Source archive directory | `C:\Archive\Originals` |
| `${MoveOcrDir}` | OCR archive directory | `C:\Archive\Processed` |
| `${MoveTxtDir}` | Text archive directory | `C:\Archive\Text` |
| `${RemoveSource}` | Source deletion flag | `True` / `False` |
| `${ForceOCR}` | Force OCR flag | `True` / `False` |
| `${EmailTo}` | Recipient addresses | `user@company.com` |
| `${EmailFiles}` | Attached file types | `Ocr, Txt` |
| `${SmtpServer}` | SMTP server | `smtp.company.local` |
| `${SmtpPort}` | SMTP port | `25` |
| `${SmtpUser}` | SMTP username | `scanner` |

**Example:**
```powershell
.\Invoke-OCR.ps1 -Path "report.pdf" -EmailTo "team@company.com" -SmtpServer "smtp.local" `
    -EmailSubject 'Scan ready: ${filename}' `
    -EmailBody 'The file "${basename}" was scanned at ${dpi} DPI using ${Language}. Took ${elapsed}s.'
```

---

## 📋 Per-Directory Configuration (.ocrconfig)

You can place a `.ocrconfig` file in any directory to automatically configure how files in that directory are processed. This works both with the background watcher and when running `Invoke-OCR.ps1` directly.

**Command-line parameters always take precedence over `.ocrconfig` values.**

### Format
Simple `Key = Value` pairs. Lines starting with `#` or `'` are comments. Values can be quoted or unquoted.

### Example `.ocrconfig`
```ini
# OCR Configuration for the finance department scans
Language = fra+eng
Dpi = 600
ForceOCR = true

# Email notifications
EmailTo = finance@company.com
EmailSubject = OCR Done: ${filename}
EmailBody = File "${filename}" processed at ${dpi} DPI with ${Language}. Elapsed: ${elapsed}s.
EmailFrom = Scanner <scanner@company.com>
EmailReplyTo = admin@company.com
SmtpServer = smtp.company.local
SmtpPort = 587

# Archive processed files
MoveSourceDir = C:\Archive\Originals
MoveOcrDir = C:\Archive\Processed
```

### Supported Keys

All `Invoke-OCR.ps1` parameters can be set in the config file:

`Language`, `Dpi`, `Page`, `ForceOCR`, `RemoveSource`, `WatermarkPdf`, `ThrottleLimit`, `MoveSourceDir`, `MoveOcrDir`, `MoveTxtDir`, `EmailTo`, `EmailSubject`, `EmailBody`, `EmailFrom`, `EmailReplyTo`, `SmtpServer`, `SmtpPort`, `SmtpUser`, `SmtpPassword`

### Hierarchical Inheritance

Config files are inherited from parent directories. A `.ocrconfig` in `C:\scans` provides defaults, and a `.ocrconfig` in `C:\scans\finance` overrides only the keys it defines:

```
C:\scans\.ocrconfig          → Base: Dpi=300, Language=eng+fra
C:\scans\finance\.ocrconfig  → Override: Language=fra, EmailTo=finance@co.com
```

Files in `C:\scans\finance\` get: `Dpi=300, Language=fra, EmailTo=finance@co.com`

---

## 🔐 Secure Credential Storage

Instead of storing SMTP passwords in plain text `.ocrconfig` files, you can use Windows DPAPI encryption:

1. **Save credentials**: Run `.\Save-OcrCredential.ps1 -Name "office365"`
2. **Reference in .ocrconfig**: `SmtpPassword = credential:office365`

The credential file is encrypted with your Windows account and cannot be read by other users or on other machines.

---

## ✅ Config Validation

Two methods to validate `.ocrconfig` files:

### Standalone Script
```powershell
# Validate a single directory
.\Test-OcrConfig.ps1 -Path "C:\scans\finance"

# Recursively validate all configs
.\Test-OcrConfig.ps1 -Path "C:\scans" -Recurse
```

### Inline Flag
```powershell
.\Invoke-OCR.ps1 -Path "C:\scans\finance" -ValidateConfig
```

Both methods check for unknown keys (with typo suggestions), duplicate keys, missing credential files, and parent config inheritance.
