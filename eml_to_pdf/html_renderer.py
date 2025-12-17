"""HTML to PDF rendering with WeasyPrint and fallback support."""

import base64
import html
import re
import logging
from typing import Dict, Optional, List
from email.message import EmailMessage

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

from .config import ConversionConfig

logger = logging.getLogger(__name__)

# Check for WeasyPrint availability
WEASYPRINT_AVAILABLE = False
_WEASYPRINT_ERROR = None

def _check_weasyprint():
    """Check if WeasyPrint is available without printing errors."""
    global WEASYPRINT_AVAILABLE, _WEASYPRINT_ERROR
    import sys
    import io
    import os

    # Suppress stdout and stderr during import
    old_stdout = sys.stdout
    old_stderr = sys.stderr
    sys.stdout = io.StringIO()
    sys.stderr = io.StringIO()

    # Also suppress WeasyPrint's internal warning system
    devnull = open(os.devnull, 'w')

    try:
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
        WEASYPRINT_AVAILABLE = True
        return True, HTML, CSS, FontConfiguration
    except ImportError:
        _WEASYPRINT_ERROR = "not installed"
        return False, None, None, None
    except OSError as e:
        _WEASYPRINT_ERROR = f"missing dependencies"
        return False, None, None, None
    except Exception as e:
        _WEASYPRINT_ERROR = str(e)
        return False, None, None, None
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr
        devnull.close()

# Perform the check
_wp_result = _check_weasyprint()
WEASYPRINT_AVAILABLE = _wp_result[0]
if WEASYPRINT_AVAILABLE:
    HTML, CSS, FontConfiguration = _wp_result[1], _wp_result[2], _wp_result[3]
    logger.debug("WeasyPrint is available for HTML rendering")
else:
    HTML, CSS, FontConfiguration = None, None, None
    logger.debug(f"WeasyPrint {_WEASYPRINT_ERROR or 'unavailable'}, using fallback HTML rendering")

# Default CSS for email rendering
EMAIL_CSS = """
@page {
    size: letter;
    margin: 0.75in;
}

body {
    font-family: Helvetica, Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.5;
    color: #333;
}

.email-header {
    border-bottom: 2px solid #ccc;
    padding-bottom: 15px;
    margin-bottom: 20px;
}

.email-header table {
    width: 100%;
    border-collapse: collapse;
}

.email-header td {
    padding: 3px 0;
    vertical-align: top;
}

.email-header .label {
    font-weight: bold;
    width: 60px;
    color: #555;
}

.email-header .value {
    word-break: break-word;
}

.email-body {
    margin-top: 20px;
}

img {
    max-width: 100%;
    height: auto;
}

a {
    color: #0066cc;
}

pre, code {
    font-family: Courier, monospace;
    background-color: #f5f5f5;
    padding: 2px 4px;
}

blockquote {
    border-left: 3px solid #ccc;
    margin-left: 0;
    padding-left: 15px;
    color: #666;
}

table {
    border-collapse: collapse;
    max-width: 100%;
}

td, th {
    border: 1px solid #ddd;
    padding: 8px;
}

.attachments-section {
    margin-top: 30px;
    padding-top: 15px;
    border-top: 1px solid #ccc;
}

.attachments-section h3 {
    margin-bottom: 10px;
    color: #555;
}

.attachments-section ul {
    list-style-type: none;
    padding-left: 0;
}

.attachments-section li {
    padding: 5px 0;
}

.att-size {
    color: #888;
    font-size: 0.9em;
}

.att-type {
    color: #666;
    font-size: 0.9em;
}
"""


def extract_cid_images(msg: EmailMessage) -> Dict[str, str]:
    """
    Extract inline images from email message parts.

    Args:
        msg: Parsed email message object

    Returns:
        Dict mapping CID references to base64 data URIs
    """
    cid_images = {}

    for part in msg.walk():
        content_type = part.get_content_type()
        content_id = part.get('Content-ID')

        if content_id and content_type.startswith('image/'):
            # Clean up Content-ID (remove < and >)
            cid = content_id.strip('<>').strip()

            # Get image data
            try:
                image_data = part.get_payload(decode=True)
                if image_data:
                    # Convert to base64 data URI
                    b64_data = base64.b64encode(image_data).decode('utf-8')
                    data_uri = f"data:{content_type};base64,{b64_data}"
                    cid_images[cid] = data_uri

                    logger.debug(f"Extracted CID image: {cid} ({len(image_data)} bytes)")
            except Exception as e:
                logger.warning(f"Failed to extract CID image {cid}: {e}")

    return cid_images


def replace_cid_references(html_content: str, cid_images: Dict[str, str]) -> str:
    """
    Replace cid: references in HTML with base64 data URIs.

    Args:
        html_content: HTML string with cid: references
        cid_images: Dict mapping CID to data URIs

    Returns:
        HTML with cid: references replaced
    """
    def replace_cid(match):
        cid = match.group(1)
        # Try exact match first
        if cid in cid_images:
            return f'src="{cid_images[cid]}"'
        # Try without any prefix/suffix
        for key, value in cid_images.items():
            if cid in key or key in cid:
                return f'src="{value}"'
        # Return placeholder if not found
        return 'src="" alt="[Image not found]"'

    # Match cid: references in src attributes
    pattern = r'src=["\']?cid:([^"\'\s>]+)["\']?'

    return re.sub(pattern, replace_cid, html_content, flags=re.IGNORECASE)


def build_email_html(
    body_html: str,
    metadata: Dict[str, str],
    embedded_images: Optional[Dict[str, str]] = None,
    attachments: Optional[List] = None,
    config: Optional[ConversionConfig] = None
) -> str:
    """
    Build a complete HTML document with email header and body.

    Args:
        body_html: The HTML body content
        metadata: Dict with subject, sender, recipients, cc, bcc, date
        embedded_images: Optional dict of CID images to embed
        attachments: Optional list of attachment info
        config: Optional configuration

    Returns:
        Complete HTML document string
    """
    config = config or ConversionConfig()

    # Process CID images if provided
    if embedded_images:
        body_html = replace_cid_references(body_html, embedded_images)

    # Build header rows based on config
    header_rows = []
    if config.include_subject:
        header_rows.append(f'<tr><td class="label">Subject:</td><td class="value">{html.escape(metadata.get("subject", "No Subject"))}</td></tr>')
    if config.include_from:
        header_rows.append(f'<tr><td class="label">From:</td><td class="value">{html.escape(metadata.get("sender", "Unknown"))}</td></tr>')
    if config.include_to:
        header_rows.append(f'<tr><td class="label">To:</td><td class="value">{html.escape(metadata.get("recipients", ""))}</td></tr>')
    if config.include_cc and metadata.get("cc") and metadata.get("cc") != "No CC":
        header_rows.append(f'<tr><td class="label">CC:</td><td class="value">{html.escape(metadata.get("cc", ""))}</td></tr>')
    if config.include_bcc and metadata.get("bcc") and metadata.get("bcc") != "No BCC":
        header_rows.append(f'<tr><td class="label">BCC:</td><td class="value">{html.escape(metadata.get("bcc", ""))}</td></tr>')
    if config.include_date:
        header_rows.append(f'<tr><td class="label">Date:</td><td class="value">{html.escape(metadata.get("date", ""))}</td></tr>')

    header_html = f"""
    <div class="email-header">
        <table>
            {''.join(header_rows)}
        </table>
    </div>
    """

    # Build attachments section
    attachments_html = ""
    if attachments:
        from .attachment_handler import format_attachment_size
        attachments_html = """
        <div class="attachments-section">
            <h3>Attachments</h3>
            <ul>
        """
        for att in attachments:
            size_str = format_attachment_size(att.size)
            attachments_html += f"""
                <li>
                    {html.escape(att.name)}
                    <span class="att-size">({size_str})</span>
                    <span class="att-type">[{html.escape(att.content_type)}]</span>
                </li>
            """
        attachments_html += """
            </ul>
        </div>
        """

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>{html.escape(metadata.get('subject', 'Email'))}</title>
    </head>
    <body>
        {header_html}
        <div class="email-body">
            {body_html}
        </div>
        {attachments_html}
    </body>
    </html>
    """


def render_html_to_pdf_weasyprint(
    html_content: str,
    output_path: str,
    metadata: Dict[str, str],
    embedded_images: Optional[Dict[str, str]] = None,
    attachments: Optional[List] = None,
    config: Optional[ConversionConfig] = None
) -> bool:
    """
    Render HTML email content to PDF using WeasyPrint.

    Args:
        html_content: The HTML body of the email
        output_path: Path for the output PDF
        metadata: Dict with subject, sender, recipients, cc, bcc, date
        embedded_images: Dict mapping CID references to base64 data URIs
        attachments: Optional list of attachment info
        config: Optional configuration

    Returns:
        True if successful, False otherwise
    """
    if not WEASYPRINT_AVAILABLE:
        return False

    config = config or ConversionConfig()

    # Adjust CSS for page size
    css_content = EMAIL_CSS
    if config.page_size.lower() == "a4":
        css_content = css_content.replace("size: letter;", "size: A4;")

    try:
        # Build complete HTML document
        full_html = build_email_html(
            html_content, metadata, embedded_images, attachments, config
        )

        font_config = FontConfiguration()
        html_doc = HTML(string=full_html)
        css = CSS(string=css_content, font_config=font_config)
        html_doc.write_pdf(output_path, stylesheets=[css], font_config=font_config)

        logger.info(f"Successfully rendered PDF with WeasyPrint: {output_path}")
        return True

    except Exception as e:
        logger.error(f"WeasyPrint rendering failed: {e}")
        return False


def extract_text_from_html(html_content: str) -> str:
    """
    Extract readable text from HTML content by cleaning up HTML tags.
    Fallback method when WeasyPrint is not available.

    Args:
        html_content: The raw HTML content from email

    Returns:
        Cleaned, readable text content
    """
    try:
        content = html_content

        # Remove scripts
        content = re.sub(
            r'<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>',
            '', content, flags=re.DOTALL
        )

        # Remove styles
        content = re.sub(
            r'<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>',
            '', content, flags=re.DOTALL
        )

        # Replace image tags with placeholders
        content = re.sub(
            r'<img[^>]*?src="cid:[^"]*"[^>]*?alt="([^"]*)"[^>]*?>',
            r'[Image: \1]', content
        )
        content = re.sub(r'<img[^>]*?alt="([^"]*)"[^>]*?>', r'[Image: \1]', content)
        content = re.sub(r'<img[^>]*?>', r'[Image]', content)

        # Replace links with text and URL
        content = re.sub(
            r'<a[^>]*?href="([^"]*)"[^>]*?>(.*?)<\/a>',
            r'\2 (\1)', content, flags=re.DOTALL
        )

        # Replace common block elements with line breaks
        for tag in ['div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'tr', 'li', 'br']:
            if tag == 'br':
                content = re.sub(f'<{tag}[^>]*?>', '\n', content)
            else:
                content = re.sub(f'</{tag}[^>]*?>', '\n', content)

        # Handle lists with bullets
        content = re.sub(r'<li[^>]*?>', '* ', content)

        # Replace table cells with spacing
        content = re.sub(r'<td[^>]*?>', ' | ', content)

        # Emphasize bold and headers
        for tag in ['b', 'strong', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            content = re.sub(f'<{tag}[^>]*?>', '*', content)
            content = re.sub(f'</{tag}>', '*', content)

        # Emphasize italic
        for tag in ['i', 'em']:
            content = re.sub(f'<{tag}[^>]*?>', '_', content)
            content = re.sub(f'</{tag}>', '_', content)

        # Remove all remaining HTML tags
        content = re.sub(r'<[^>]*?>', '', content)

        # Decode HTML entities
        content = html.unescape(content)

        # Fix multiple line breaks
        content = re.sub(r'\n\s*\n', '\n\n', content)

        # Trim extra whitespace
        content = re.sub(r'[ \t]+', ' ', content)

        return content.strip()

    except Exception as e:
        logger.error(f"Error extracting text from HTML: {e}")
        # Return minimal cleaned content on error
        return re.sub(r'<[^>]*?>', ' ', html_content)


def render_html_to_pdf_reportlab(
    html_content: str,
    output_path: str,
    metadata: Dict[str, str],
    config: Optional[ConversionConfig] = None
) -> bool:
    """
    Create a PDF from email HTML content using ReportLab.
    Fallback method when WeasyPrint is not available.

    Args:
        html_content: The HTML content to convert
        output_path: The output PDF path
        metadata: Dictionary containing email metadata
        config: Optional configuration

    Returns:
        True if successful, False if failed
    """
    config = config or ConversionConfig()

    try:
        # Extract readable text from HTML
        extracted_text = extract_text_from_html(html_content)

        # Create PDF with ReportLab
        doc = SimpleDocTemplate(output_path, pagesize=config.get_page_size())
        styles = getSampleStyleSheet()

        # Create a Title style with more space
        title_style = styles["Title"]

        # Add metadata in a structured format
        elements = []

        # Add email subject as title
        if config.include_subject:
            subject = metadata.get('subject', 'No Subject')
            elements.append(Paragraph(html.escape(subject), title_style))
            elements.append(Spacer(1, 20))

        # Add metadata table
        metadata_parts = []
        if config.include_from:
            metadata_parts.append(f"<b>From:</b> {html.escape(metadata.get('sender', 'Unknown Sender'))}")
        if config.include_to:
            metadata_parts.append(f"<b>To:</b> {html.escape(metadata.get('recipients', 'No Recipients'))}")
        if config.include_cc and metadata.get('cc') and metadata.get('cc') != 'No CC':
            metadata_parts.append(f"<b>CC:</b> {html.escape(metadata.get('cc', ''))}")
        if config.include_bcc and metadata.get('bcc') and metadata.get('bcc') != 'No BCC':
            metadata_parts.append(f"<b>BCC:</b> {html.escape(metadata.get('bcc', ''))}")
        if config.include_date:
            metadata_parts.append(f"<b>Date:</b> {html.escape(metadata.get('date', 'Unknown Date'))}")

        if metadata_parts:
            metadata_text = "<br/>".join(metadata_parts)
            elements.append(Paragraph(metadata_text, styles["Normal"]))
            elements.append(Spacer(1, 20))

        # Add a divider
        elements.append(Paragraph("<hr/>", styles["Normal"]))
        elements.append(Spacer(1, 10))

        # Add note about HTML content
        elements.append(Paragraph(
            "<i>Note: This email contained HTML content. Formatting has been simplified for compatibility.</i>",
            styles["Normal"]
        ))
        elements.append(Spacer(1, 20))

        # Process the extracted text
        formatted_text = html.escape(extracted_text).replace('\n', '<br/>')

        try:
            elements.append(Paragraph(formatted_text, styles["Normal"]))
        except Exception as parser_error:
            logger.error(f"Error parsing HTML with ReportLab: {parser_error}")
            safe_text = html.escape(extracted_text).replace('\n', '<br/>')
            elements.append(Paragraph(
                f"<i>Note: Additional formatting issues were detected.</i><br/><br/>{safe_text}",
                styles["Normal"]
            ))

        doc.build(elements)
        return True

    except Exception as e:
        logger.error(f"Error converting HTML to PDF with ReportLab: {e}")
        return False


def render_html_to_pdf(
    html_content: str,
    output_path: str,
    metadata: Dict[str, str],
    embedded_images: Optional[Dict[str, str]] = None,
    attachments: Optional[List] = None,
    config: Optional[ConversionConfig] = None
) -> bool:
    """
    Render HTML email content to PDF with automatic fallback.

    Tries WeasyPrint first if available and configured, falls back to ReportLab.

    Args:
        html_content: The HTML body of the email
        output_path: Path for the output PDF
        metadata: Dict with subject, sender, recipients, cc, bcc, date
        embedded_images: Dict mapping CID references to base64 data URIs
        attachments: Optional list of attachment info
        config: Optional configuration

    Returns:
        True if successful, False otherwise
    """
    config = config or ConversionConfig()

    # Try WeasyPrint first if available and configured
    if WEASYPRINT_AVAILABLE and config.use_weasyprint:
        success = render_html_to_pdf_weasyprint(
            html_content, output_path, metadata, embedded_images, attachments, config
        )
        if success:
            return True
        logger.warning("WeasyPrint failed, falling back to ReportLab")

    # Fallback to ReportLab
    return render_html_to_pdf_reportlab(html_content, output_path, metadata, config)
