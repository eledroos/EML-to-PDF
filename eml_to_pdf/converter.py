"""Core email to PDF conversion logic."""

import os
import html
import logging
from email import policy
from email.parser import BytesParser
from email.message import EmailMessage
from typing import List, Dict, Optional, Tuple, Set, Callable
from dataclasses import dataclass

from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

from .config import ConversionConfig
from .utils import format_filename, get_year_month_from_date, logger
from .html_renderer import (
    render_html_to_pdf,
    extract_cid_images,
    WEASYPRINT_AVAILABLE
)
from .attachment_handler import extract_attachments, AttachmentInfo
from .contact_extractor import extract_contacts_from_eml, generate_address_book, Contact


@dataclass
class ConversionResult:
    """Result of a single email conversion."""
    success: bool
    source_file: str
    output_path: Optional[str] = None
    error_message: Optional[str] = None
    attachments: Optional[List[AttachmentInfo]] = None


@dataclass
class BatchConversionResult:
    """Result of a batch conversion operation."""
    total_files: int
    successful: int
    failed: int
    results: List[ConversionResult]
    output_folder: str
    cancelled: bool = False
    address_book_path: Optional[str] = None


def extract_metadata(msg: EmailMessage) -> Dict[str, str]:
    """
    Extract metadata from an email message.

    Args:
        msg: Parsed email message

    Returns:
        Dict with subject, sender, recipients, cc, bcc, date
    """
    return {
        'subject': str(msg['subject'] or "No Subject"),
        'sender': str(msg['from'] or "Unknown Sender"),
        'recipients': str(msg['to'] or "No Recipients"),
        'cc': str(msg['cc'] or "No CC"),
        'bcc': str(msg['bcc'] or "No BCC"),
        'date': str(msg['date'] or "Unknown Date"),
    }


def convert_plaintext_to_pdf(
    text_content: str,
    output_path: str,
    metadata: Dict[str, str],
    config: Optional[ConversionConfig] = None
) -> bool:
    """
    Convert plain text email content to PDF.

    Args:
        text_content: The plain text content
        output_path: Path for the output PDF
        metadata: Dict with email metadata
        config: Optional configuration

    Returns:
        True if successful, False otherwise
    """
    config = config or ConversionConfig()

    try:
        doc = SimpleDocTemplate(output_path, pagesize=config.get_page_size())
        styles = getSampleStyleSheet()

        elements = []

        # Add metadata
        if config.include_subject:
            elements.append(Paragraph(
                f"<b>Subject:</b> {html.escape(metadata.get('subject', ''))}",
                styles["Normal"]
            ))
        if config.include_from:
            elements.append(Paragraph(
                f"<b>From:</b> {html.escape(metadata.get('sender', ''))}",
                styles["Normal"]
            ))
        if config.include_to:
            elements.append(Paragraph(
                f"<b>To:</b> {html.escape(metadata.get('recipients', ''))}",
                styles["Normal"]
            ))
        if config.include_cc and metadata.get('cc') and metadata.get('cc') != 'No CC':
            elements.append(Paragraph(
                f"<b>CC:</b> {html.escape(metadata.get('cc', ''))}",
                styles["Normal"]
            ))
        if config.include_bcc and metadata.get('bcc') and metadata.get('bcc') != 'No BCC':
            elements.append(Paragraph(
                f"<b>BCC:</b> {html.escape(metadata.get('bcc', ''))}",
                styles["Normal"]
            ))
        if config.include_date:
            elements.append(Paragraph(
                f"<b>Date:</b> {html.escape(metadata.get('date', ''))}",
                styles["Normal"]
            ))

        elements.append(Spacer(1, 20))

        # Add body
        safe_body = html.escape(text_content).replace("\n", "<br />")
        elements.append(Paragraph(f"<b>Body:</b><br />{safe_body}", styles["Normal"]))

        doc.build(elements)
        return True

    except Exception as e:
        logger.error(f"Error converting plain text to PDF: {e}")
        return False


def convert_single_email(
    eml_path: str,
    output_folder: str,
    processed_files: Set[str],
    config: Optional[ConversionConfig] = None
) -> ConversionResult:
    """
    Convert a single EML file to PDF.

    Args:
        eml_path: Path to the EML file
        output_folder: Base output folder
        processed_files: Set of already used filenames (for uniqueness)
        config: Optional configuration

    Returns:
        ConversionResult with success status and details
    """
    config = config or ConversionConfig()
    eml_file = os.path.basename(eml_path)

    try:
        # Parse email
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        # Extract metadata
        metadata = extract_metadata(msg)

        # Get year/month for folder organization
        date_str = metadata.get('date', '')
        year, month = get_year_month_from_date(date_str)

        # Determine output folder
        if config.organize_by_date:
            final_output_folder = os.path.join(output_folder, year, month)
        else:
            final_output_folder = output_folder

        os.makedirs(final_output_folder, exist_ok=True)

        # Get email body
        body_part = msg.get_body(preferencelist=('plain', 'html'))
        if not body_part:
            return ConversionResult(
                success=False,
                source_file=eml_file,
                error_message="No plain text or HTML body found"
            )

        # Format filename
        new_filename = format_filename(date_str, metadata['subject'], processed_files)
        processed_files.add(new_filename)

        # Create PDF path
        pdf_path = os.path.join(final_output_folder, f"{new_filename}.pdf")

        # Extract attachments if configured
        attachments = None
        if config.extract_attachments:
            attachments = extract_attachments(msg, final_output_folder, new_filename)

        # Handle different content types
        if body_part.get_content_type() == 'text/html':
            html_content = body_part.get_content()

            # Extract CID images for embedding
            cid_images = extract_cid_images(msg)

            success = render_html_to_pdf(
                html_content,
                pdf_path,
                metadata,
                embedded_images=cid_images,
                attachments=attachments,
                config=config
            )

            if not success:
                return ConversionResult(
                    success=False,
                    source_file=eml_file,
                    error_message="Failed to convert HTML content to PDF"
                )
        else:
            # Plain text content
            body = body_part.get_content()
            success = convert_plaintext_to_pdf(body, pdf_path, metadata, config)

            if not success:
                return ConversionResult(
                    success=False,
                    source_file=eml_file,
                    error_message="Failed to convert plain text to PDF"
                )

        return ConversionResult(
            success=True,
            source_file=eml_file,
            output_path=pdf_path,
            attachments=attachments
        )

    except Exception as e:
        logger.error(f"Error converting {eml_file}: {e}")
        return ConversionResult(
            success=False,
            source_file=eml_file,
            error_message=str(e)
        )


def convert_batch(
    input_folder: str,
    output_folder: Optional[str] = None,
    config: Optional[ConversionConfig] = None,
    progress_callback: Optional[Callable[[int, int, str], bool]] = None
) -> BatchConversionResult:
    """
    Convert all EML files in a folder to PDF.

    Args:
        input_folder: Folder containing EML files
        output_folder: Output folder (default: input_folder/PDF)
        config: Optional configuration
        progress_callback: Optional callback(current, total, filename) -> continue
                          Return False to cancel

    Returns:
        BatchConversionResult with statistics and details
    """
    config = config or ConversionConfig()

    # Setup output folder
    if output_folder is None:
        output_folder = os.path.join(input_folder, "PDF")
    os.makedirs(output_folder, exist_ok=True)

    # Find EML files
    eml_files = [
        f for f in os.listdir(input_folder)
        if f.lower().endswith(".eml")
    ]

    if not eml_files:
        return BatchConversionResult(
            total_files=0,
            successful=0,
            failed=0,
            results=[],
            output_folder=output_folder
        )

    # Sort by modification time (oldest first)
    eml_files.sort(key=lambda f: os.path.getmtime(os.path.join(input_folder, f)))

    results = []
    processed_files: Set[str] = set()
    all_contacts: List[Contact] = []
    successful = 0
    failed = 0
    cancelled = False

    for i, eml_file in enumerate(eml_files):
        # Check for cancellation via callback
        if progress_callback:
            should_continue = progress_callback(i, len(eml_files), eml_file)
            if not should_continue:
                cancelled = True
                break

        eml_path = os.path.join(input_folder, eml_file)
        result = convert_single_email(eml_path, output_folder, processed_files, config)
        results.append(result)

        if result.success:
            successful += 1

        # Extract contacts if address book generation is enabled
        if config.generate_address_book:
            contacts = extract_contacts_from_eml(eml_path)
            all_contacts.extend(contacts)

        if not result.success:
            failed += 1

    # Final progress callback
    if progress_callback and not cancelled:
        progress_callback(len(eml_files), len(eml_files), "Complete")

    # Generate address book if enabled and we have contacts
    address_book_path = None
    if config.generate_address_book and all_contacts and not cancelled:
        from pathlib import Path
        csv_path = Path(output_folder) / "address_book.csv"
        result_path = generate_address_book(all_contacts, csv_path, dedupe=True)
        if result_path:
            address_book_path = str(result_path)

    return BatchConversionResult(
        total_files=len(eml_files),
        successful=successful,
        failed=failed,
        results=results,
        output_folder=output_folder,
        cancelled=cancelled,
        address_book_path=address_book_path
    )


def create_skipped_files_report(
    results: List[ConversionResult],
    output_folder: str
) -> Optional[str]:
    """
    Create a PDF report of skipped/failed files.

    Args:
        results: List of conversion results
        output_folder: Folder to save the report

    Returns:
        Path to the report PDF, or None if no failures
    """
    failed_results = [r for r in results if not r.success]

    if not failed_results:
        return None

    report_path = os.path.join(output_folder, "Skipped_Files_Report.pdf")

    try:
        from reportlab.lib.pagesizes import letter

        doc = SimpleDocTemplate(report_path, pagesize=letter)
        styles = getSampleStyleSheet()

        elements = [
            Paragraph("<b>Skipped Files Report</b>", styles["Title"]),
            Spacer(1, 20),
            Paragraph(
                "<b>The following files were skipped during processing:</b>",
                styles["Normal"]
            ),
            Spacer(1, 10)
        ]

        for result in failed_results:
            safe_entry = html.escape(
                f"{result.source_file}: {result.error_message or 'Unknown error'}"
            )
            elements.append(Paragraph(safe_entry, styles["Normal"]))
            elements.append(Spacer(1, 5))

        doc.build(elements)
        return report_path

    except Exception as e:
        logger.error(f"Error creating skipped files report: {e}")
        return None
