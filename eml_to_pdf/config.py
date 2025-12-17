"""Configuration management for EML to PDF Converter."""

import json
import os
from dataclasses import dataclass, field, asdict
from typing import List, Optional
from pathlib import Path

# Default config file location
CONFIG_PATH = Path.home() / ".eml_to_pdf_config.json"


@dataclass
class ConversionConfig:
    """Configuration options for EML to PDF conversion."""

    # Page settings
    page_size: str = "letter"  # "letter" or "a4"
    font_family: str = "Helvetica"
    font_size: int = 11

    # Output settings
    organize_by_date: bool = True  # Organize output into YYYY/MM folders

    # Metadata fields to include
    include_subject: bool = True
    include_from: bool = True
    include_to: bool = True
    include_cc: bool = True
    include_bcc: bool = True
    include_date: bool = True

    # Processing options
    extract_attachments: bool = False
    attachment_folder: str = "attachments"
    use_weasyprint: bool = True  # Use WeasyPrint for HTML rendering

    # UI settings
    theme: str = "cosmo"  # ttkbootstrap theme name

    @classmethod
    def load(cls, path: Optional[Path] = None) -> "ConversionConfig":
        """
        Load configuration from file.

        Args:
            path: Optional path to config file. Uses default if not specified.

        Returns:
            ConversionConfig instance
        """
        config_path = path or CONFIG_PATH

        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    data = json.load(f)
                return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})
            except (json.JSONDecodeError, TypeError) as e:
                # Return defaults if config is corrupted
                return cls()

        return cls()

    def save(self, path: Optional[Path] = None) -> None:
        """
        Save configuration to file.

        Args:
            path: Optional path to config file. Uses default if not specified.
        """
        config_path = path or CONFIG_PATH

        with open(config_path, 'w') as f:
            json.dump(asdict(self), f, indent=2)

    def get_metadata_fields(self) -> List[str]:
        """Get list of enabled metadata fields."""
        fields = []
        if self.include_subject:
            fields.append('subject')
        if self.include_from:
            fields.append('from')
        if self.include_to:
            fields.append('to')
        if self.include_cc:
            fields.append('cc')
        if self.include_bcc:
            fields.append('bcc')
        if self.include_date:
            fields.append('date')
        return fields

    def get_page_size(self):
        """Get reportlab page size tuple."""
        from reportlab.lib.pagesizes import letter, A4

        if self.page_size.lower() == "a4":
            return A4
        return letter


# Available themes for the UI
AVAILABLE_THEMES = [
    "cosmo",      # Light, clean
    "flatly",     # Light, flat design
    "litera",     # Light, subtle
    "minty",      # Light, green accents
    "lumen",      # Light, minimal
    "sandstone",  # Light, warm
    "yeti",       # Light, rounded
    "pulse",      # Light, purple accents
    "united",     # Light, orange accents
    "morph",      # Light, morphing effects
    "journal",    # Light, journal-like
    "darkly",     # Dark theme
    "superhero",  # Dark, blue accents
    "solar",      # Dark, warm
    "cyborg",     # Dark, futuristic
    "vapor",      # Dark, neon
]

# Available page sizes
PAGE_SIZES = ["letter", "a4"]

# Available fonts (common PDF fonts)
AVAILABLE_FONTS = [
    "Helvetica",
    "Times-Roman",
    "Courier",
]
