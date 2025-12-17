"""Shared utility functions for EML to PDF conversion."""

import os
import re
import logging
from datetime import datetime
from typing import Set, Optional

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def format_filename(date: str, subject: str, existing_names: Set[str]) -> str:
    """
    Format the filename as 'DATE - Subject' and ensure uniqueness.

    Args:
        date: Email date string
        subject: Email subject
        existing_names: Set of already used filenames

    Returns:
        Unique formatted filename (without extension)
    """
    safe_subject = sanitize_filename(subject[:50])

    try:
        formatted_date = datetime.strptime(
            date, "%a, %d %b %Y %H:%M:%S %z"
        ).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        # Try alternative date formats
        for fmt in [
            "%d %b %Y %H:%M:%S %z",
            "%Y-%m-%d %H:%M:%S",
            "%a, %d %b %Y %H:%M:%S",
        ]:
            try:
                formatted_date = datetime.strptime(date, fmt).strftime("%Y-%m-%d")
                break
            except (ValueError, TypeError):
                continue
        else:
            formatted_date = "Unknown_Date"

    base_name = f"{formatted_date} - {safe_subject}".strip()

    # Ensure unique name
    unique_name = base_name
    counter = 1
    while unique_name in existing_names:
        unique_name = f"{base_name} ({counter})"
        counter += 1

    return unique_name


def sanitize_filename(filename: str) -> str:
    """
    Sanitize a string for use as a filename.

    Args:
        filename: The string to sanitize

    Returns:
        A safe filename string
    """
    # Remove or replace dangerous characters
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
    # Replace multiple underscores/spaces with single
    safe = re.sub(r'[_\s]+', ' ', safe)
    # Remove leading/trailing whitespace
    safe = safe.strip()
    # Limit length
    if len(safe) > 100:
        safe = safe[:100]
    return safe if safe else "untitled"


def get_unique_filepath(folder: str, filename: str) -> str:
    """
    Get a unique filepath, adding numbers if file exists.

    Args:
        folder: The target directory
        filename: The desired filename

    Returns:
        A unique filepath
    """
    filepath = os.path.join(folder, filename)
    if not os.path.exists(filepath):
        return filepath

    name, ext = os.path.splitext(filename)
    counter = 1
    while os.path.exists(filepath):
        filepath = os.path.join(folder, f"{name}_{counter}{ext}")
        counter += 1

    return filepath


def parse_email_date(date_str: str) -> Optional[datetime]:
    """
    Parse an email date string into a datetime object.

    Args:
        date_str: The date string from an email header

    Returns:
        datetime object or None if parsing fails
    """
    if not date_str:
        return None

    # Common email date formats
    formats = [
        "%a, %d %b %Y %H:%M:%S %z",
        "%d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S",
        "%d %b %Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S %z",
        "%Y-%m-%d %H:%M:%S",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except (ValueError, TypeError):
            continue

    return None


def get_year_month_from_date(date_str: str) -> tuple:
    """
    Extract year and month from an email date string.

    Args:
        date_str: The date string from an email header

    Returns:
        Tuple of (year, month) strings
    """
    parsed = parse_email_date(date_str)
    if parsed:
        return parsed.strftime("%Y"), parsed.strftime("%m")
    return "Unknown_Year", "Unknown_Month"
