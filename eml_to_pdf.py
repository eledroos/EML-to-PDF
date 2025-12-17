#!/usr/bin/env python3
"""
EML to PDF Converter

A tool for converting EML email files to PDF format.

This is a backward-compatible wrapper script. For the full feature set,
use the module directly:
    python -m eml_to_pdf [options]

Usage:
    python eml_to_pdf.py          # Launch GUI
    python eml_to_pdf.py -i ./emails  # Convert from command line

Author: Nasser Eledroos
License: CC0 (Public Domain)
"""

import sys

# Add the current directory to the path to allow importing the package
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from eml_to_pdf import main

if __name__ == "__main__":
    sys.exit(main())
