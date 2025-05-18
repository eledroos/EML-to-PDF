EML to PDF Converter
====================

A simple Python-based tool to convert `.eml` files (email files) into beautifully formatted `.pdf` files. This tool is designed to handle batch processing with a progress bar and includes features like skipped file reporting and customizable PDF formatting.

---

Features
--------
- Convert multiple `.eml` files to PDF in a single batch.
- Extracts and includes metadata (Subject, Sender, Recipients, CC, BCC, Date) in the PDFs.
- Handles word wrapping and long email content gracefully.
- Saves a detailed report (`Skipped_Files_Report.pdf`) for any skipped files.
- User-friendly GUI with progress bar and status messages.
- Fully open-source under CC0 (Public Domain).

---

Requirements
------------
- Python 3.7+
- Required dependency: `pip install reportlab`

The tool handles both plain text and HTML emails, converting them to formatted PDFs. HTML content is converted to readable text with formatting cues where possible.

---

Usage
-----
1. Clone this repository.
2. Run the script: `python eml_to_pdf.py`
3. Follow the on-screen instructions to select a folder containing `.eml` files. The converted PDFs will be saved in a `PDF` subfolder.

---

macOS Gatekeeper Bypass Instructions
-----

This app is not signed with an Apple Developer ID. When you open the app for the first time, macOS will show a security warning. Follow these steps to run the app:

1. Download and unzip the app:
    - Click the download link for the .zip file.
    - Double-click the .zip file to extract the app.

2. Right-click to Open:
    - Right-click (or Control-click) the app and select Open.
    - A warning will appear saying, "The app cannot be opened because it is from an unidentified developer."

3. Confirm and Open:
    - In the warning dialog, click Open to confirm.
    - The app will now run.

4. Run Normally in the Future:
    - After the first successful run, you can open the app by double-clicking it as usual.

License
-------
This project is released under the **CC0 Public Domain Dedication**. See the `LICENSE` file for details.

---

Author
------
Created by **Nasser Eledroos**. Contributions and feedback are welcome!