import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path
from pdf2docx import Converter
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import subprocess
import threading
import os
import sys

# --- Working directory relative to script location ---
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "Output"
OUTPUT_DIR.mkdir(exist_ok=True)

# --- Check LibreOffice ---
LIBREOFFICE_PATH = os.environ.get("LIBREOFFICE_PATH") or r"C:\Program Files\LibreOffice\program\soffice.exe"
if not Path(LIBREOFFICE_PATH).exists():
    messagebox.showerror("Error",
                         f"LibreOffice not found at {LIBREOFFICE_PATH}. Please install it or set the environment variable LIBREOFFICE_PATH.")
    sys.exit(1)


# --- Conversion functions ---
def pdf_to_word(file_path):
    word_file = OUTPUT_DIR / f"{file_path.stem}_converted.docx"
    if not file_path.exists():
        raise FileNotFoundError(f"{file_path} does not exist")
    with Converter(str(file_path)) as cv:
        cv.convert(str(word_file), start=0, end=None)

    doc = Document(str(word_file))
    for para in doc.paragraphs:
        if any('\u0600' <= c <= '\u06FF' for c in para.text):  # Arabic text
            para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    doc.save(str(word_file))
    return word_file


def word_to_pdf(file_path):
    pdf_file = OUTPUT_DIR / f"{file_path.stem}_converted.pdf"
    if not file_path.exists():
        raise FileNotFoundError(f"{file_path} does not exist")
    subprocess.run([
        LIBREOFFICE_PATH,
        "--headless",
        "--convert-to", "pdf",
        str(file_path),
        "--outdir", str(OUTPUT_DIR)
    ], check=True)
    return pdf_file


# --- Convert all files ---
def convert_all_files():
    files = list(BASE_DIR.glob("*.pdf")) + list(BASE_DIR.glob("*.docx"))
    if not files:
        messagebox.showwarning("Warning", "No PDF or Word files found in the folder.")
        return

    progress['maximum'] = len(files)

    for i, file_path in enumerate(files, start=1):
        try:
            if file_path.suffix.lower() == ".pdf":
                pdf_to_word(file_path)
            elif file_path.suffix.lower() == ".docx":
                word_to_pdf(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Error converting {file_path.name}:\n{e}")
        progress['value'] = i
        root.update_idletasks()

    messagebox.showinfo("Done", f"All files have been converted successfully to the 'Output' folder.")


# --- GUI ---
root = tk.Tk()
root.title("PDF ↔ Word Converter Auto")
root.geometry("600x250")

label = tk.Label(root, text="All PDF and Word files in the folder will be converted automatically.", font=("Arial", 12))
label.pack(pady=20)

btn_convert_all = tk.Button(root, text="Start Conversion", font=("Arial", 12), bg="green", fg="white",
                            command=lambda: threading.Thread(target=convert_all_files, daemon=True).start())
btn_convert_all.pack(pady=10)

progress = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
progress.pack(pady=20)

root.mainloop()
