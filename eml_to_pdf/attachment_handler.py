"""Attachment extraction and handling for EML files."""

import os
import re
import mimetypes
import logging
from email.message import EmailMessage
from typing import List, Dict, Optional
from dataclasses import dataclass

from .utils import sanitize_filename, get_unique_filepath

logger = logging.getLogger(__name__)


@dataclass
class AttachmentInfo:
    """Information about an extracted attachment."""
    name: str
    path: str
    size: int
    content_type: str


def extract_attachments(
    msg: EmailMessage,
    output_folder: str,
    pdf_filename: str
) -> List[AttachmentInfo]:
    """
    Extract attachments from an email message.

    Args:
        msg: Parsed email message object
        output_folder: Base folder for output
        pdf_filename: Name of the PDF file (without extension) for folder naming

    Returns:
        List of AttachmentInfo objects for extracted attachments
    """
    attachments = []
    attachment_folder = os.path.join(output_folder, f"{pdf_filename}_attachments")
    attachment_count = 0

    for part in msg.walk():
        # Skip multipart containers
        if part.get_content_maintype() == 'multipart':
            continue

        # Get content disposition and ID
        content_disposition = part.get('Content-Disposition', '')
        content_id = part.get('Content-ID')
        content_type = part.get_content_type()

        # Determine if this is an attachment
        is_attachment = False

        if 'attachment' in content_disposition.lower():
            is_attachment = True
        elif part.get_filename():
            # Has filename - check if it's not an inline image
            if 'inline' not in content_disposition.lower() or not content_id:
                is_attachment = True

        if not is_attachment:
            continue

        # Get filename
        filename = part.get_filename()
        if not filename:
            # Generate filename from content type
            ext = guess_extension(content_type) or '.bin'
            attachment_count += 1
            filename = f"attachment_{attachment_count}{ext}"

        # Sanitize filename
        safe_filename = sanitize_filename(filename)
        if not safe_filename:
            safe_filename = f"attachment_{attachment_count}"

        # Get attachment data
        try:
            data = part.get_payload(decode=True)
            if data is None:
                continue

            # Create attachments folder if needed
            os.makedirs(attachment_folder, exist_ok=True)

            # Handle duplicate filenames
            filepath = get_unique_filepath(attachment_folder, safe_filename)

            # Write attachment
            with open(filepath, 'wb') as f:
                f.write(data)

            attachments.append(AttachmentInfo(
                name=filename,
                path=filepath,
                size=len(data),
                content_type=content_type
            ))

            logger.info(f"Extracted attachment: {filename} ({len(data)} bytes)")

        except Exception as e:
            logger.error(f"Failed to extract attachment {filename}: {e}")

    return attachments


def guess_extension(content_type: str) -> Optional[str]:
    """
    Guess file extension from content type.

    Args:
        content_type: MIME content type

    Returns:
        File extension including the dot, or None
    """
    # Use mimetypes library
    ext = mimetypes.guess_extension(content_type)
    if ext:
        return ext

    # Common fallbacks
    type_to_ext = {
        'application/pdf': '.pdf',
        'application/msword': '.doc',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
        'application/vnd.ms-excel': '.xls',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
        'application/vnd.ms-powerpoint': '.ppt',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
        'application/zip': '.zip',
        'application/x-rar-compressed': '.rar',
        'application/x-7z-compressed': '.7z',
        'text/plain': '.txt',
        'text/html': '.html',
        'text/csv': '.csv',
        'image/jpeg': '.jpg',
        'image/png': '.png',
        'image/gif': '.gif',
        'image/webp': '.webp',
        'image/svg+xml': '.svg',
        'audio/mpeg': '.mp3',
        'audio/wav': '.wav',
        'video/mp4': '.mp4',
        'video/webm': '.webm',
    }

    return type_to_ext.get(content_type.lower())


def format_attachment_size(size_bytes: int) -> str:
    """
    Format attachment size in human-readable format.

    Args:
        size_bytes: Size in bytes

    Returns:
        Formatted size string (e.g., "1.5 MB")
    """
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"
