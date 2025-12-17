# EML to PDF Converter

A powerful Python tool to convert `.eml` email files into formatted PDF documents. Features batch processing, HTML email rendering with embedded images, attachment extraction, and a modern GUI with drag-and-drop support.

## Features

- **Batch Conversion** - Convert multiple `.eml` files in a single operation
- **HTML Email Rendering** - Properly renders HTML emails with formatting (using WeasyPrint)
- **Embedded Images** - Extracts and embeds inline CID images in PDFs
- **Attachment Extraction** - Optionally saves email attachments alongside PDFs
- **CLI & GUI** - Use from command line for automation or via graphical interface
- **Drag-and-Drop** - Drop folders directly onto the GUI to convert
- **Organized Output** - PDFs automatically organized by year/month
- **Progress Tracking** - Visual progress with ETA and cancel support
- **Settings Panel** - Customize page size, fonts, metadata fields
- **Skipped Files Report** - Detailed PDF report for any failed conversions
- **Cross-Platform** - Works on macOS, Windows, and Linux

## Installation

### From Source (Recommended)

```bash
# Clone the repository
git clone https://github.com/nassereledroos/EML-to-PDF.git
cd EML-to-PDF

# Install dependencies
pip install -r requirements.txt

# Run the application
python eml_to_pdf.py
```

### Optional Dependencies

For the full feature set, install all optional dependencies:

```bash
pip install weasyprint ttkbootstrap tkinterdnd2
```

- **weasyprint** - Better HTML email rendering with CSS support
- **ttkbootstrap** - Modern themed GUI
- **tkinterdnd2** - Drag-and-drop support

## Usage

### GUI Mode (Default)

Simply run without arguments to launch the graphical interface:

```bash
python eml_to_pdf.py
# or
python -m eml_to_pdf
```

### Command Line

```bash
# Convert emails from a folder
python -m eml_to_pdf -i ~/emails

# Specify output folder
python -m eml_to_pdf -i ~/emails -o ~/pdfs

# Extract attachments
python -m eml_to_pdf -i ~/emails --extract-attachments

# Flat output (no year/month folders)
python -m eml_to_pdf -i ~/emails --no-organize

# Show help
python -m eml_to_pdf --help
```

### CLI Options

| Option | Description |
|--------|-------------|
| `-i, --input FOLDER` | Input folder containing EML files |
| `-o, --output FOLDER` | Output folder for PDFs (default: INPUT/PDF) |
| `--page-size {letter,a4}` | PDF page size (default: letter) |
| `--extract-attachments` | Extract and save email attachments |
| `--no-organize` | Don't organize by year/month folders |
| `--no-weasyprint` | Disable WeasyPrint (use simple text rendering) |
| `-v, --verbose` | Verbose output |
| `-q, --quiet` | Suppress output except errors |
| `--gui` | Force GUI mode |

## Building Standalone Executables

Build standalone applications that don't require Python:

```bash
# Install build dependencies
pip install -r requirements-dev.txt

# Build for current platform
python build.py

# Build and package for release
python build.py --package

# Clean build artifacts
python build.py --clean
```

### Platform-Specific Scripts

```bash
# macOS
./scripts/build_macos.sh

# Windows
scripts\build_windows.bat

# Linux
./scripts/build_linux.sh
```

### Automated Releases

Push a version tag to trigger GitHub Actions builds for all platforms:

```bash
git tag v2.0.0
git push origin v2.0.0
```

## Project Structure

```
EML-to-PDF/
├── eml_to_pdf.py           # Entry point (wrapper script)
├── eml_to_pdf/             # Main package
│   ├── __init__.py
│   ├── __main__.py         # Module entry point
│   ├── cli.py              # Command-line interface
│   ├── gui.py              # Graphical interface
│   ├── converter.py        # Core conversion logic
│   ├── html_renderer.py    # HTML to PDF rendering
│   ├── attachment_handler.py
│   ├── config.py           # Settings management
│   └── utils.py            # Shared utilities
├── build.py                # Build script
├── requirements.txt        # Runtime dependencies
├── requirements-dev.txt    # Build dependencies
├── pyproject.toml          # Project metadata
├── assets/                 # Application icons
├── scripts/                # Platform build scripts
└── .github/workflows/      # CI/CD workflows
```

## Configuration

Settings are stored in `~/.eml_to_pdf_config.json` and can be configured via the GUI Settings panel:

- **Page Size** - Letter or A4
- **Font** - Helvetica, Times-Roman, or Courier
- **Output Structure** - Organize by year/month or flat
- **Metadata Fields** - Choose which fields to include
- **Attachments** - Enable/disable extraction
- **HTML Rendering** - Use WeasyPrint or simple text

## macOS Gatekeeper Notice

Downloaded executables are not signed with an Apple Developer ID. On first launch:

1. Right-click (or Control-click) the app
2. Select **Open** from the context menu
3. Click **Open** in the security dialog
4. Future launches will work normally

## Requirements

- Python 3.8+
- reportlab (required)
- weasyprint (optional, for HTML rendering)
- ttkbootstrap (optional, for modern UI)
- tkinterdnd2 (optional, for drag-and-drop)

## License

This project is released under the **CC0 Public Domain Dedication**. See the [LICENSE](LICENSE) file for details.

## Author

Created by **Nasser Eledroos**

Contributions and feedback are welcome!
