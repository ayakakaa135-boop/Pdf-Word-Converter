# PDF ↔ Word Converter

A Python application to **automatically convert PDF files to Word and Word files to PDF**.  
It also handles **Arabic text alignment** and provides a simple **GUI**.

## Features

- Automatically converts all PDF and Word files in the project folder.
- Converts **PDF → Word** and **Word → PDF**.
- Adjusts **Arabic text alignment** in Word automatically.
- Saves all converted files in an **Output** folder.
- Shows **progress bar** during conversion.
- GUI built with **Tkinter**, responsive during processing.

## Requirements

- Python 3.10+  
- Libraries:
  - `pdf2docx`
  - `python-docx`
  - `tkinter` (usually included with Python)
- **LibreOffice** installed for Word → PDF conversion (Windows):
  ```
  C:\Program Files\LibreOffice\program\soffice.exe
  ```

## Installation

1. Clone or download this project.
2. Open a terminal in the project folder.
3. Install required libraries:
   ```bash
   pip install pdf2docx python-docx
   ```

## Usage

1. Place all PDF and Word files you want to convert inside the project folder:  
   ```
   C:\Users\haama\Desktop\PycharmProjects\Advanced Topics in Python\pdf2word
   ```
2. Run the main Python script:
   ```bash
   python main.py
   ```
3. Click the **"Start Conversion"** button in the GUI.
4. Wait for all files to be processed. Converted files will appear in the **Output** folder.

## Notes

- PDF → Word conversion keeps text and paragraphs; images are converted approximately.
- Word → PDF conversion uses LibreOffice, no Microsoft Word required.
- Arabic text alignment is adjusted automatically in Word.

## License

This project is free to use and modify.

