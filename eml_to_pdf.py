import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from email import policy
from email.parser import BytesParser
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer


def display_help():
    """Display the help information."""
    help_message = (
        "What is an EML file?\n"
        "EML files store email messages saved from email clients like Outlook or Thunderbird.\n\n\n"
        "What does this software do?\n"
        "This tool converts EML files into PDF format, embedding email metadata (sender, recipient, subject, "
        "CC, BCC, date) and the body of the email into the PDF.\n\n\n"
        "Who made this?\n"
        "Created by Nasser Eledroos\n"
        "License: CC0 (No rights reserved, public domain)."
    )
    messagebox.showinfo("Help - About EML to PDF Converter", help_message)


def format_filename(date, subject, existing_names):
    """Format the filename as 'DATE - Subject' and ensure uniqueness."""
    safe_subject = "".join(c if c.isalnum() or c in " _-." else "_" for c in subject[:50])  # Limit length for filenames
    formatted_date = datetime.strptime(date, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")
    base_name = f"{formatted_date} - {safe_subject}".strip()

    # Ensure unique name
    unique_name = base_name
    counter = 1
    while unique_name in existing_names:
        unique_name = f"{base_name} ({counter})"
        counter += 1

    return unique_name


def convert_eml_to_pdf():
    # Ask user for a folder with EML files
    eml_folder = filedialog.askdirectory(title="Select Folder Containing EML Files")
    if not eml_folder:
        messagebox.showerror("No Folder Selected", "Please select a folder containing EML files.")
        return

    # Prepare output folder
    output_folder = os.path.join(eml_folder, "PDF")
    os.makedirs(output_folder, exist_ok=True)

    # Gather EML files
    eml_files = [f for f in os.listdir(eml_folder) if f.lower().endswith(".eml")]
    if not eml_files:
        messagebox.showerror("No EML Files Found", "The selected folder contains no EML files.")
        return

    # Setup progress bar
    progress["value"] = 0
    progress["maximum"] = len(eml_files)

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

            # Extract body content safely
            if not msg.get_body(preferencelist=('plain')):  # Check if body exists
                skipped_files.append(eml_file)  # Add to skipped files
                continue

            body = msg.get_body(preferencelist=('plain')).get_content()

            # Format the filename
            new_filename = format_filename(date, subject, processed_files)
            processed_files.add(new_filename)

            # Create PDF
            pdf_path = os.path.join(output_folder, f"{new_filename}.pdf")
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
            formatted_body = body.replace("\n", "<br />")
            elements.append(Paragraph(f"<b>Body:</b><br />{formatted_body}", styles["Normal"]))

            doc.build(elements)

        except Exception as e:
            print(f"Failed to process {eml_file}: {e}")
            skipped_files.append(eml_file)  # Add problematic file to skipped list
            continue

        # Update progress bar
        progress["value"] = i + 1
        app.update_idletasks()

    # Create a Skipped Files PDF if there are skipped files
    if skipped_files:
        skipped_report_path = os.path.join(output_folder, "Skipped_Files_Report.pdf")
        doc = SimpleDocTemplate(skipped_report_path, pagesize=letter)
        styles = getSampleStyleSheet()
        elements = [Paragraph("<b>Skipped Files Report</b>", styles["Title"])]
        elements.append(Spacer(1, 20))
        for file in skipped_files:
            elements.append(Paragraph(file, styles["Normal"]))
        doc.build(elements)

    # Show a pretty final message
    if skipped_files:
        final_message = (
            "✅ Conversion Complete! 🎉\n\n"
            f"Converted: {len(processed_files)} EML files to PDF.\n"
            f"Location: {output_folder}\n\n"
            f"⚠️ Skipped: {len(skipped_files)} files.\n"
            f"A detailed list has been saved in 'Skipped_Files_Report.pdf' in the same folder."
        )
    else:
        final_message = (
            "✅ Conversion Complete!\n\n"
            f"Converted: {len(processed_files)} EML files to PDF.\n"
            f"Location: {output_folder}\n\n"
            "🎉 No files were skipped!"
        )

    messagebox.showinfo("Conversion Complete", final_message)


# Create GUI
app = tk.Tk()
app.title("EML to PDF Converter")
app.geometry("500x350")

# GUI components
label = tk.Label(
    app,
    text=(
        "Welcome to the EML to PDF Converter!\n"
        "Click 'Convert EML Files' to select a folder with EML files.\n"
        "Converted PDFs will be saved in a 'PDF' subfolder."
    ),
    font=("Arial", 12),
    wraplength=400,
    justify="center"
)
label.pack(pady=20)

convert_button = tk.Button(app, text="Convert EML Files", command=convert_eml_to_pdf, font=("Arial", 14))
convert_button.pack(pady=10)

help_button = tk.Button(app, text="Help", command=display_help, font=("Arial", 14))
help_button.pack(pady=10)

progress = ttk.Progressbar(app, orient="horizontal", length=400, mode="determinate")
progress.pack(pady=20)

# Start the Tkinter event loop
app.mainloop()
