"""Contact extraction and address book generation."""

import csv
import re
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, getaddresses
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .utils import logger


@dataclass
class Contact:
    """Represents an email contact."""
    name: str
    email: str
    contact_type: str  # from, to, cc, bcc


def parse_email_header(header_value: str) -> List[Tuple[str, str]]:
    """
    Parse an email header value into (name, email) tuples.

    Handles formats like:
    - "john@example.com"
    - "John Doe <john@example.com>"
    - "John Doe <john@example.com>, Jane <jane@example.com>"

    Args:
        header_value: Raw header value string

    Returns:
        List of (name, email) tuples
    """
    if not header_value or header_value in ("No Recipients", "No CC", "No BCC", "Unknown Sender"):
        return []

    # Use email.utils.getaddresses for robust parsing
    addresses = getaddresses([header_value])

    result = []
    for name, email in addresses:
        if email:  # Only include if there's an actual email
            # Clean up name - use email prefix if no name provided
            if not name:
                name = email.split('@')[0] if '@' in email else email
            result.append((name.strip(), email.strip().lower()))

    return result


def extract_contacts_from_metadata(metadata: Dict[str, str], source_file: str = "") -> List[Contact]:
    """
    Extract all contacts from email metadata.

    Args:
        metadata: Dict with 'sender', 'recipients', 'cc', 'bcc' keys
        source_file: Optional source filename for reference

    Returns:
        List of Contact objects
    """
    contacts = []

    # Map metadata keys to contact types
    field_mapping = [
        ('sender', 'From'),
        ('recipients', 'To'),
        ('cc', 'CC'),
        ('bcc', 'BCC'),
    ]

    for field_key, contact_type in field_mapping:
        header_value = metadata.get(field_key, '')
        parsed = parse_email_header(header_value)

        for name, email in parsed:
            contacts.append(Contact(
                name=name,
                email=email,
                contact_type=contact_type
            ))

    return contacts


def extract_contacts_from_eml(eml_path: Path) -> List[Contact]:
    """
    Extract all contacts from an EML file.

    Args:
        eml_path: Path to the EML file

    Returns:
        List of Contact objects
    """
    try:
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        metadata = {
            'sender': str(msg['from'] or ''),
            'recipients': str(msg['to'] or ''),
            'cc': str(msg['cc'] or ''),
            'bcc': str(msg['bcc'] or ''),
        }

        return extract_contacts_from_metadata(metadata)

    except Exception as e:
        logger.error(f"Error extracting contacts from {eml_path}: {e}")
        return []


def deduplicate_contacts(contacts: List[Contact]) -> List[Contact]:
    """
    Remove duplicate contacts by email address.

    Keeps the first occurrence of each email.

    Args:
        contacts: List of Contact objects

    Returns:
        Deduplicated list of contacts
    """
    seen_emails = set()
    unique_contacts = []

    for contact in contacts:
        email_lower = contact.email.lower()
        if email_lower not in seen_emails:
            seen_emails.add(email_lower)
            unique_contacts.append(contact)

    return unique_contacts


def generate_address_book(
    contacts: List[Contact],
    output_path: Path,
    dedupe: bool = True
) -> Optional[Path]:
    """
    Generate a CSV address book from contacts.

    Args:
        contacts: List of Contact objects
        output_path: Path for the output CSV file
        dedupe: Whether to deduplicate by email address

    Returns:
        Path to the generated CSV, or None if no contacts
    """
    if not contacts:
        logger.info("No contacts to export")
        return None

    # Deduplicate if requested
    if dedupe:
        contacts = deduplicate_contacts(contacts)

    # Sort by name for easier reading
    contacts.sort(key=lambda c: (c.name.lower(), c.email.lower()))

    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)

            # Write header
            writer.writerow(['Name', 'Email', 'Type'])

            # Write contacts
            for contact in contacts:
                writer.writerow([contact.name, contact.email, contact.contact_type])

        logger.info(f"Address book saved: {output_path} ({len(contacts)} contacts)")
        return output_path

    except Exception as e:
        logger.error(f"Error generating address book: {e}")
        return None
