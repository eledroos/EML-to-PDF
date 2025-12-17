"""
EML to PDF Converter

A tool for converting EML email files to PDF format with support for
HTML rendering, inline images, and attachments.

Usage:
    # As a module
    python -m eml_to_pdf --input ./emails --output ./pdfs

    # Launch GUI
    python -m eml_to_pdf --gui

    # Or simply
    python -m eml_to_pdf

Author: Nasser Eledroos
License: CC0 (Public Domain)
"""

__version__ = "2.0.0"
__author__ = "Nasser Eledroos"

from .config import ConversionConfig
from .converter import (
    convert_batch,
    convert_single_email,
    ConversionResult,
    BatchConversionResult,
)
from .cli import main
from .gui import launch_gui

__all__ = [
    "ConversionConfig",
    "convert_batch",
    "convert_single_email",
    "ConversionResult",
    "BatchConversionResult",
    "main",
    "launch_gui",
    "__version__",
]
