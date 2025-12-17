"""
Entry point for running EML to PDF Converter as a module.

Usage:
    python -m eml_to_pdf [options]
    python -m eml_to_pdf --gui
    python -m eml_to_pdf -i ./emails -o ./pdfs
"""

import sys
from .cli import main

if __name__ == "__main__":
    sys.exit(main())
