# Word to PDF Converter App

This repository contains a small Tkinter-based application that converts `.docx` files to PDF using the [docx2pdf](https://pypi.org/project/docx2pdf/) library.

## Requirements
- Python 3.8+
- The `docx2pdf` package (install with `pip install docx2pdf`)
- On Windows, Microsoft Word is typically required.
- On Linux/macOS, `LibreOffice` may be used if Word is not available.

## Running the Application
```bash
python word_to_pdf_converter.py
```

1. Click **"Browse for .docx Files"** and select one or more Word documents.
2. (Optional) Choose an output directory for the PDFs. If not specified, PDFs are saved alongside the original files.
3. Click **"Convert Selected Files to PDF"** to perform the conversion.
4. Once finished, click **"Finish & Exit"** to close the program.

