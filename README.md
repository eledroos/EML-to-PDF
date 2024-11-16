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
- Install dependencies using: `pip install reportlab`

---

Usage
-----
1. Clone this repository.
2. Run the script: `python eml_to_pdf.py`
3. Follow the on-screen instructions to select a folder containing `.eml` files. The converted PDFs will be saved in a `PDF` subfolder.

---

License
-------
This project is released under the **CC0 Public Domain Dedication**. See the `LICENSE` file for details.

---

Author
------
Created by **Nasser Eledroos**. Contributions and feedback are welcome!