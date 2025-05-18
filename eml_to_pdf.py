import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from email import policy
from email.parser import BytesParser
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.platypus.flowables import Image
import tempfile
import base64
import re
import os
import html
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Check for xhtml2pdf but don't use it - it doesn't handle CID image references correctly
try:
    from xhtml2pdf import pisa
    XHTML2PDF_AVAILABLE = False  # Disable for now even if available, as it's causing issues
    logger.info("xhtml2pdf found but disabled for compatibility with email CID references")
except ImportError:
    XHTML2PDF_AVAILABLE = False
    logger.info("xhtml2pdf not found, using simplified HTML processing")


def display_help():
    """Display the help information."""
    help_message = (
        "What is an EML file?\n"
        "EML files store email messages saved from email clients like Outlook or Thunderbird.\n\n"
        "What does this software do?\n"
        "This tool converts EML files into PDF format, embedding email metadata (Subject, Sender, Recipients, etc.) "
        "and the body of the email into the PDF.\n\n"
        "Who made this?\n"
        "Created by Nasser Eledroos\n"
        "License: CC0 (Public Domain)."
    )
    messagebox.showinfo("Help - About EML to PDF Converter", help_message)


def format_filename(date, subject, existing_names):
    """Format the filename as 'DATE - Subject' and ensure uniqueness."""
    safe_subject = "".join(c if c.isalnum() or c in " _-." else "_" for c in subject[:50])  # Limit length for filenames
    try:
        formatted_date = datetime.strptime(date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")
    except ValueError:
        formatted_date = "Unknown_Date"

    base_name = f"{formatted_date} - {safe_subject}".strip()

    # Ensure unique name
    unique_name = base_name
    counter = 1
    while unique_name in existing_names:
        unique_name = f"{base_name} ({counter})"
        counter += 1

    return unique_name


def extract_text_from_html(html_content):
    """
    Extract readable text from HTML content by cleaning up HTML tags.
    Handles complex HTML with CID images better than direct conversion.
    
    Args:
        html_content: The raw HTML content from email
        
    Returns:
        Cleaned, readable text content
    """
    try:
        # Remove scripts
        content = re.sub(r'<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>', '', html_content, flags=re.DOTALL)
        
        # Remove styles
        content = re.sub(r'<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>', '', content, flags=re.DOTALL)
        
        # Replace image tags with placeholders to avoid parsing issues
        content = re.sub(r'<img[^>]*?src="cid:[^"]*"[^>]*?alt="([^"]*)"[^>]*?>', r'[Image: \1]', content)
        content = re.sub(r'<img[^>]*?alt="([^"]*)"[^>]*?>', r'[Image: \1]', content)
        content = re.sub(r'<img[^>]*?>', r'[Image]', content)
        
        # Replace links with text and URL
        content = re.sub(r'<a[^>]*?href="([^"]*)"[^>]*?>(.*?)<\/a>', r'\2 (\1)', content, flags=re.DOTALL)
        
        # Replace common block elements with line breaks
        for tag in ['div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'tr', 'li', 'br']:
            if tag == 'br':
                pattern = f'<{tag}[^>]*?>'
                content = re.sub(pattern, '\n', content)
            else:
                pattern = f'</{tag}[^>]*?>'
                content = re.sub(pattern, '\n', content)
        
        # Handle lists with bullets
        content = re.sub(r'<li[^>]*?>', r'• ', content)
        
        # Replace table cells with spacing
        content = re.sub(r'<td[^>]*?>', r' | ', content)
        
        # Emphasize bold and headers
        for tag in ['b', 'strong', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            pattern_start = f'<{tag}[^>]*?>'
            pattern_end = f'</{tag}>'
            content = re.sub(pattern_start, '*', content)
            content = re.sub(pattern_end, '*', content)
        
        # Emphasize italic
        for tag in ['i', 'em']:
            pattern_start = f'<{tag}[^>]*?>'
            pattern_end = f'</{tag}>'
            content = re.sub(pattern_start, '_', content)
            content = re.sub(pattern_end, '_', content)
        
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


def convert_html_to_pdf(html_content, output_path, metadata):
    """
    Create a PDF from email HTML content.
    Uses a reliable text extraction approach that works with any email format.
    
    Args:
        html_content: The HTML content to convert
        output_path: The output PDF path
        metadata: Dictionary containing email metadata
    
    Returns:
        True if successful, False if failed
    """
    try:
        # Extract readable text from HTML
        extracted_text = extract_text_from_html(html_content)
        
        # Create PDF with ReportLab
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        styles = getSampleStyleSheet()
        
        # Create a Title style with more space
        title_style = styles["Title"]
        
        # Add metadata in a structured format
        elements = []
        
        # Add email subject as title
        subject = metadata.get('subject', 'No Subject')
        elements.append(Paragraph(subject, title_style))
        elements.append(Spacer(1, 20))
        
        # Add metadata table
        metadata_text = f"""
        <b>From:</b> {metadata.get('sender', 'Unknown Sender')}<br/>
        <b>To:</b> {metadata.get('recipients', 'No Recipients')}<br/>
        <b>CC:</b> {metadata.get('cc', 'No CC')}<br/>
        <b>BCC:</b> {metadata.get('bcc', 'No BCC')}<br/>
        <b>Date:</b> {metadata.get('date', 'Unknown Date')}<br/>
        """
        elements.append(Paragraph(metadata_text, styles["Normal"]))
        elements.append(Spacer(1, 20))
        
        # Add a divider
        elements.append(Paragraph("<hr/>", styles["Normal"]))
        elements.append(Spacer(1, 10))
        
        # Add note about HTML content
        elements.append(Paragraph("<i>Note: This email contained HTML content. Formatting has been simplified for compatibility.</i>", styles["Normal"]))
        elements.append(Spacer(1, 20))
        
        # Process the extracted text, replacing newlines with <br/>
        formatted_text = extracted_text.replace('\n', '<br/>')
        
        # Wrap this in a try-except in case there's any ReportLab parsing issues
        try:
            elements.append(Paragraph(formatted_text, styles["Normal"]))
        except Exception as parser_error:
            logger.error(f"Error parsing HTML with ReportLab: {parser_error}")
            # Fallback to extreme sanitization if even the extracted text has issues
            safe_text = html.escape(extracted_text).replace('\n', '<br/>')
            elements.append(Paragraph(f"<i>Note: Additional formatting issues were detected.</i><br/><br/>{safe_text}", styles["Normal"]))
        
        doc.build(elements)
        return True
        
    except Exception as e:
        logger.error(f"Error converting HTML to PDF: {e}")
        
        # Last resort fallback - create a minimal PDF with just metadata and error notice
        try:
            doc = SimpleDocTemplate(output_path, pagesize=letter)
            styles = getSampleStyleSheet()
            
            elements = []
            subject = metadata.get('subject', 'No Subject')
            elements.append(Paragraph(subject, styles["Title"]))
            elements.append(Spacer(1, 20))
            
            # Add metadata
            for key, value in metadata.items():
                if value and value.lower() not in ['no cc', 'no bcc', 'no recipients', 'unknown sender', 'unknown date']:
                    elements.append(Paragraph(f"<b>{key.capitalize()}:</b> {value}", styles["Normal"]))
            
            elements.append(Spacer(1, 20))
            elements.append(Paragraph("<b>Note:</b> This email contained complex HTML content that could not be fully converted. "
                               "The content may be better viewed in an email client.", styles["Normal"]))
            
            doc.build(elements)
            return True
        except Exception as fallback_error:
            logger.error(f"Critical error in fallback PDF creation: {fallback_error}")
            return False


def create_progress_popup():
    """Create a pop-up window for the progress bar."""
    popup = tk.Toplevel(app)
    popup.title("Processing...")
    popup.geometry("400x100")
    popup.resizable(False, False)
    label = tk.Label(popup, text="Processing files, please wait...", font=("Arial", 12))
    label.pack(pady=10)
    progress = ttk.Progressbar(popup, orient="horizontal", length=300, mode="determinate")
    progress.pack(pady=10)
    return popup, progress


def convert_eml_to_pdf():
    # Ask user for a folder with EML files
    eml_folder = filedialog.askdirectory(title="Select Folder Containing EML Files")
    if not eml_folder:
        messagebox.showerror("No Folder Selected", "Please select a folder containing EML files.")
        return

    # Prepare base output folder
    base_output_folder = os.path.join(eml_folder, "PDF")
    os.makedirs(base_output_folder, exist_ok=True)

    # Gather EML files
    eml_files = [f for f in os.listdir(eml_folder) if f.lower().endswith(".eml")]
    if not eml_files:
        messagebox.showerror("No EML Files Found", "The selected folder contains no EML files.")
        return

    # Create the progress bar pop-up
    popup, progress = create_progress_popup()
    progress["maximum"] = len(eml_files)
    popup.update()

    skipped_files = []  # Track skipped files
    processed_files = set()  # Track processed files to ensure unique filenames

    for i, eml_file in enumerate(sorted(eml_files, key=lambda f: os.path.getmtime(os.path.join(eml_folder, f)))):  # Oldest first
        try:
            eml_path = os.path.join(eml_folder, eml_file)
            with open(eml_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)

            # Extract metadata
            subject = msg['subject'] or "No Subject"
            sender = msg['from'] or "Unknown Sender"
            recipients = msg['to'] or "No Recipients"
            cc = msg['cc'] or "No CC"
            bcc = msg['bcc'] or "No BCC"
            date = msg['date'] or "Unknown Date"

            # Parse the date to organize by year/month
            try:
                email_date = datetime.strptime(date, "%a, %d %b %Y %H:%M:%S %z")
                year = email_date.strftime("%Y")
                month = email_date.strftime("%m")
            except ValueError:
                year = "Unknown_Year"
                month = "Unknown_Month"

            # Create year/month subfolders
            output_folder = os.path.join(base_output_folder, year, month)
            os.makedirs(output_folder, exist_ok=True)

            # Extract body content safely - try plain text first, then HTML if plain not available
            body_part = msg.get_body(preferencelist=('plain', 'html'))
            if not body_part:  # Check if any body exists
                skipped_reason = f"{eml_file}: No plain text or HTML body found"
                skipped_files.append(skipped_reason)
                continue

            # Format the filename
            new_filename = format_filename(date, subject, processed_files)
            processed_files.add(new_filename)

            # Create PDF path
            pdf_path = os.path.join(output_folder, f"{new_filename}.pdf")
            
            # Create metadata dictionary
            metadata = {
                'subject': subject,
                'sender': sender,
                'recipients': recipients,
                'cc': cc,
                'bcc': bcc,
                'date': date
            }
            
            # Handle different content types
            if body_part.get_content_type() == 'text/html':
                html_content = body_part.get_content()
                
                # Process HTML content with our specialized function
                success = convert_html_to_pdf(html_content, pdf_path, metadata)
                
                if not success:
                    skipped_reason = f"{eml_file}: Failed to convert HTML content to PDF"
                    skipped_files.append(skipped_reason)
                    continue
            else:
                # Text content - use regular ReportLab approach
                body = body_part.get_content()
                doc = SimpleDocTemplate(pdf_path, pagesize=letter)
                styles = getSampleStyleSheet()

                # Build PDF content
                elements = []
                elements.append(Paragraph(f"<b>Subject:</b> {subject}", styles["Normal"]))
                elements.append(Paragraph(f"<b>From:</b> {sender}", styles["Normal"]))
                elements.append(Paragraph(f"<b>To:</b> {recipients}", styles["Normal"]))
                elements.append(Paragraph(f"<b>CC:</b> {cc}", styles["Normal"]))
                elements.append(Paragraph(f"<b>BCC:</b> {bcc}", styles["Normal"]))
                elements.append(Paragraph(f"<b>Date:</b> {date}", styles["Normal"]))
                elements.append(Spacer(1, 20))
                # Safely escape and format plain text
                import html
                safe_body = html.escape(body).replace("\n", "<br />")
                elements.append(Paragraph(f"<b>Body:</b><br />{safe_body}", styles["Normal"]))
                doc.build(elements)

        except Exception as e:
            # Add problematic file to skipped list with error reason
            skipped_reason = f"{eml_file}: {str(e)}"
            skipped_files.append(skipped_reason)
            continue

        # Update progress bar
        progress["value"] = i + 1
        popup.update()

    # Close the progress bar pop-up
    popup.destroy()

    # Create a Skipped Files PDF if there are skipped files
    if skipped_files:
        skipped_report_path = os.path.join(base_output_folder, "Skipped_Files_Report.pdf")
        doc = SimpleDocTemplate(skipped_report_path, pagesize=letter)
        styles = getSampleStyleSheet()
        elements = [Paragraph("<b>Skipped Files Report</b>", styles["Title"])]
        elements.append(Spacer(1, 20))
        elements.append(Paragraph("<b>The following files were skipped during processing with reasons:</b>", styles["Normal"]))
        elements.append(Spacer(1, 10))
        for entry in skipped_files:
            # Escape any HTML/XML special characters in the error message
            import html
            safe_entry = html.escape(entry)
            elements.append(Paragraph(safe_entry, styles["Normal"]))
            elements.append(Spacer(1, 5))
        doc.build(elements)

    # Show a pretty final message
    final_message = (
        f"✅ Conversion Complete!\n\n"
        f"Converted: {len(processed_files)} EML files to PDF.\n"
        f"Location: {base_output_folder}\n\n"
    )
    if skipped_files:
        final_message += (
            f"⚠️ Skipped: {len(skipped_files)} files.\n"
            f"A detailed list has been saved in 'Skipped_Files_Report.pdf' in the same folder."
        )
    else:
        final_message += "🎉 No files were skipped!"

    messagebox.showinfo(
        "Conversion Complete",
        final_message,
    )

# Create the main GUI
app = tk.Tk()
app.title("EML to PDF Converter")
app.geometry("500x350")
app.resizable(False, False)  # Disable resizing for a fixed layout

# Style Configuration
style = ttk.Style()
style.configure("TButton", font=("Arial", 14), padding=10)  # Button styling
style.configure("TLabel", font=("Arial", 12), padding=5)  # General label styling
style.configure("Title.TLabel", font=("Arial", 16, "bold"), anchor="center")  # Title styling

# Title Label
title_label = ttk.Label(
    app,
    text="EML to PDF Converter",
    style="Title.TLabel"
)
title_label.pack(pady=15)

# Instruction Label
instruction_label = ttk.Label(
    app,
    text=(
        "Welcome to the EML to PDF Converter!\n\n"
        "Click 'Convert EML Files' to select a folder with .eml files.\n"
        "Converted PDFs will be saved in a 'PDF' subfolder organized by year and month."
    ),
    style="TLabel",
    wraplength=450,
    justify="center"
)
instruction_label.pack(pady=10)

# Convert Button
convert_button = ttk.Button(
    app,
    text="Convert EML Files",
    command=convert_eml_to_pdf
)
convert_button.pack(pady=15)

# Help Button
help_button = ttk.Button(
    app,
    text="Help",
    command=display_help
)
help_button.pack(pady=5)

# Start the Tkinter event loop
app.mainloop()

