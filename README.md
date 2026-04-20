# Invoke-OCR 🔍

A robust, intelligent PowerShell automation script designed to seamlessly extract, analyze, and convert static documents (PDFs and images) into fully searchable OCR'd PDFs.

## 🌟 Key Features

- **Intelligent Pre-flight Checks**: Uses `Ghostscript` to invisibly scan PDFs to see if they already contain a text-layer. If they do, it saves you processing time by aggressively skipping them (unless bypassed via `-ForceOCR`).
- **Dynamic Path Discovery**: No need to hardcode paths. The script natively hunts for your executables checking your user arguments, scanning your system `$env:PATH`, and recursively digging through your `C:\Program Files` looking for your missing binaries.
- **Timestamp Logic**: Avoids re-processing the same files over and over. If an `_ocr.pdf` already exists and is newer than the source, the script gracefully moves onto the next task.
- **Fail-safe Error Logging**: A failure doesn't crash the loop. If a specific page corrupts during Ghostscript bursting or Tesseract parsing, the script catches it, isolates a `document.err.log` file, and keeps working through your pipeline. 
- **Bulletproof Parsing**: Fully compatible with `-LiteralPath` so filenames with spaces, hyphens (`-`), brackets (`[]`), or foreign accents never break your shell!

## ⚙️ Prerequisites

This script acts as the orchestration layer between three powerful open-source utilities. You must have these installed on your Windows machine:

1. **[Ghostscript](https://ghostscript.com/releases/gsdnld.html)**: Used for exploding static PDFs into high-resolution images, and detecting existing text strings.
2. **[Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)**: The core Optical Character Recognition engine.
3. **[PDFtk Server](https://www.pdflabs.com/tools/pdftk-server/)**: Used to rapidly safely merge the resulting standalone OCR'd pages back into a cohesive, single PDF.

## 🚀 Usage

The script is built natively for PowerShell Advanced Functions, meaning it seamlessly accepts pipeline inputs and parameters.

### Basic Run
Run standard OCR on a single document defaults (6 languages, 300 DPI):
```powershell
.\Invoke-OCR.ps1 -Path "C:\My Scans\document.pdf"
```

### Automation Pipeline
Process an entire folder of PNGs quietly, automatically answering "YES" to the executable prompts:
```powershell
Get-ChildItem -Filter "*.png" | .\Invoke-OCR.ps1 -Quiet -y
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
| `-y` | `Switch` | `$false` | Unattended bypass. Disables the interactive user-prompts confirming executables. |
| `-Quiet` | `Switch` | `$false` | Console verbosity lock. Suppresses success/progress messages but allows Warning & Errors. |
| `-Silent` | `Switch` | `$false` | The nuclear option. Mutes everything. Designed for background scheduled tasks. |
| `-ForceOCR` | `Switch` | `$false` | Unlocks files that fail the `txtwrite` pre-flight check, processing them regardless of existing text. |

*You can also directly inject overriding paths via `-TesseractPath`, `-GhostscriptPath`, and `-PdftkPath`.*
