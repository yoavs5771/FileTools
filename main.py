import os
import re
import shutil
import time
import threading
import subprocess
from tkinterdnd2 import TkinterDnD, DND_FILES
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from pathlib import Path
from docx2pdf import convert as convert_docx_to_pdf
from docx import Document
from pptx import Presentation
from comtypes.client import CreateObject
from PyPDF2 import PdfMerger
from PIL import Image

# =================== Utility Functions ===================

def extract_number(filename):
    """Extract a number from the filename (used for sorting)."""
    match = re.search(r'(\d+)', filename)
    return int(match.group(1)) if match else float('inf')

def convert_docx_to_txt(docx_path, txt_path):
    """Convert DOCX to plain TXT."""
    doc = Document(docx_path)
    with open(txt_path, "w", encoding="utf-8") as f:
        for para in doc.paragraphs:
            f.write(para.text + "\n")

def convert_docx_to_html(docx_path, html_path):
    """Convert DOCX to basic HTML format."""
    doc = Document(docx_path)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write("<html><body>\n")
        for para in doc.paragraphs:
            f.write(f"<p>{para.text}</p>\n")
        f.write("</body></html>")

def convert_pptx_to_pdf(pptx_path, pdf_path):
    """Convert PPTX to PDF using PowerPoint COM interface (Windows only)."""
    powerpoint = CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    try:
        deck = powerpoint.Presentations.Open(str(pptx_path), WithWindow=False)
        deck.SaveAs(str(pdf_path), 32)  # 32 = PDF format
        deck.Close()
    finally:
        powerpoint.Quit()

def open_folder(path):
    """Open a folder in the system file explorer."""
    if os.name == 'nt':
        os.startfile(str(path))
    elif os.name == 'posix':
        subprocess.run(["xdg-open", str(path)])

# =================== Localization Dictionary ===================

texts_dict = {
    "en": {
        "select_files": "Select files to convert",
        "select_folder": "Select folder to convert",
        "select_output": "Select output folder",
        "converting": "Converting... please wait.",
        "done": "✅ Conversion complete!",
        "fill_fields": "Please fill all fields before converting.",
        "error": "Missing Info",
        "merge_title": "Merged PDF",
        "merge_prompt": "Enter name for merged PDF (without extension):",
        "cancel": "Canceled",
        "no_merge_name": "No name provided for merged PDF. Canceling."
    },
    "he": {
        "select_files": "בחר קבצים להמרה",
        "select_folder": "בחר תיקייה להמרה",
        "select_output": "בחר תיקיית יעד",
        "converting": "ממיר... אנא המתן",
        "done": "✅ ההמרה הסתיימה!",
        "fill_fields": "אנא מלא את כל השדות לפני התחלת המרה.",
        "error": "חסר מידע",
        "merge_title": "קובץ PDF ממוזג",
        "merge_prompt": "הזן שם לקובץ הממוזג (ללא סיומת):",
        "cancel": "בוטל",
        "no_merge_name": "לא הוזן שם לקובץ הממוזג. הפעולה בוטלה."
    }
}

def set_language(lang):
    """Set the GUI language from dictionary."""
    global texts
    texts = texts_dict[lang]
    lang_var.set(lang)

# =================== GUI Actions ===================

def browse_input():
    """Open file/folder selector depending on mode."""
    if input_choice.get() == "files":
        files = filedialog.askopenfilenames(title=texts["select_files"])
        input_var.set("\n".join(files))
    else:
        folder = filedialog.askdirectory(title=texts["select_folder"])
        input_var.set(folder)

def browse_output():
    """Open folder selector for output location."""
    folder = filedialog.askdirectory(title=texts["select_output"])
    output_var.set(folder)

def handle_drop(event):
    """Handle drag & drop file or folder input."""
    paths = root.tk.splitlist(event.data)
    clean = [Path(p.strip('{}')) for p in paths if Path(p.strip('{}')).exists()]
    if clean:
        if clean[0].is_dir():
            input_choice.set("folder")
            input_var.set(str(clean[0]))
        else:
            input_choice.set("files")
            input_var.set("\n".join(str(p) for p in clean))

def threaded_conversion():
    """Prepare GUI interaction before background conversion."""
    if merge_var.get() and format_vars["pdf"].get():
        merged_name = simpledialog.askstring(texts["merge_title"], texts["merge_prompt"])
        if not merged_name:
            messagebox.showwarning(texts["cancel"], texts["no_merge_name"])
            return
    else:
        merged_name = None

    threading.Thread(target=start_conversion, args=(merged_name,), daemon=True).start()

def start_conversion(merged_name=None):
    """Run the actual file conversion."""
    log = []
    progress['value'] = 0
    progress_label.config(text=texts["converting"])
    root.update_idletasks()

    inputs = input_var.get().strip()
    output = Path(output_var.get().strip())
    selected_formats = [fmt for fmt, var in format_vars.items() if var.get()]
    merge = merge_var.get()

    if not inputs or not output or not selected_formats:
        messagebox.showerror(texts["error"], texts["fill_fields"])
        return

    input_paths = []
    if '\n' in inputs:
        input_paths = [Path(p) for p in inputs.split('\n') if Path(p).exists()]
    else:
        p = Path(inputs)
        input_paths = list(p.glob("*")) if p.is_dir() else [p]

    image_types = {"jpg", "jpeg", "png", "bmp", "gif", "tiff"}
    pdfs_created = []

    total_steps = len(input_paths) * len(selected_formats)
    progress['maximum'] = total_steps
    step = 0

    for file in input_paths:
        ext = file.suffix.lower().lstrip('.')
        stem = file.stem

        for fmt in selected_formats:
            out_dir = output / f"{fmt}_output"
            out_dir.mkdir(exist_ok=True)
            save_path = out_dir / f"{stem}.{fmt}"
            counter = 1
            while save_path.exists():
                save_path = out_dir / f"{stem}_{counter}.{fmt}"
                counter += 1

            try:
                if ext == "docx":
                    if fmt == "pdf":
                        convert_docx_to_pdf(str(file), str(save_path))
                        pdfs_created.append(save_path)
                    elif fmt == "txt":
                        convert_docx_to_txt(file, save_path)
                    elif fmt == "html":
                        convert_docx_to_html(file, save_path)

                elif ext == "pptx" and fmt == "pdf":
                    convert_pptx_to_pdf(file, save_path)
                    time.sleep(1)
                    pdfs_created.append(save_path)

                elif ext == "pdf" and fmt == "pdf":
                    shutil.copy(file, save_path)
                    pdfs_created.append(save_path)

                elif ext in image_types and fmt in image_types and fmt != ext:
                    with Image.open(file) as img:
                        rgb_img = img.convert("RGB") if img.mode in ("RGBA", "P") else img
                        rgb_img.save(save_path, fmt.upper())

                log.append(f"[OK] {file.name} -> {save_path.name}")
            except Exception as e:
                log.append(f"[FAIL] {file.name} to {fmt.upper()} – {e}")

            step += 1
            progress['value'] = step
            root.update_idletasks()

    if merge and "pdf" in selected_formats:
        try:
            merger = PdfMerger()
            for pdf in sorted(pdfs_created, key=lambda f: extract_number(f.name)):
                merger.append(str(pdf))
            merged_path = output / f"{merged_name}.pdf"
            counter = 1
            while merged_path.exists():
                merged_path = output / f"{merged_name}_{counter}.pdf"
                counter += 1
            merger.write(str(merged_path))
            merger.close()
            log.append(f"[MERGED] PDF saved to: {merged_path}")
        except Exception as e:
            log.append(f"[ERROR] Merging failed: {e}")

    with open(output / "conversion_log.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(log))

    progress_label.config(text=texts["done"])
    open_folder(output)

# =================== GUI Setup ===================

root = TkinterDnD.Tk()
root.title("🛠️ File Converter Tool")
root.geometry("650x720")

# Language selection
lang_var = tk.StringVar(value="en")
tk.Label(root, text="🌐 Language / שפה:").pack(anchor="w", padx=10)
tk.Radiobutton(root, text="English", variable=lang_var, value="en", command=lambda: set_language("en")).pack(anchor="w", padx=20)
tk.Radiobutton(root, text="עברית", variable=lang_var, value="he", command=lambda: set_language("he")).pack(anchor="w", padx=20)
set_language("en")

# Input/Output selectors
input_choice = tk.StringVar(value="files")
input_var = tk.StringVar()
output_var = tk.StringVar()
merge_var = tk.BooleanVar()
format_vars = {fmt: tk.BooleanVar() for fmt in ["pdf", "txt", "html", "jpg", "png", "bmp", "tiff"]}

tk.Label(root, text="Choose input type:").pack(anchor="w", padx=10, pady=(10, 0))
tk.Radiobutton(root, text="Select Files", variable=input_choice, value="files").pack(anchor="w", padx=20)
tk.Radiobutton(root, text="Select Folder", variable=input_choice, value="folder").pack(anchor="w", padx=20)

tk.Button(root, text="📂 Browse Input", command=browse_input).pack(pady=5)
tk.Entry(root, textvariable=input_var, width=70).pack(padx=10, pady=5)

tk.Label(root, text="Choose output folder:").pack(anchor="w", padx=10, pady=(10, 0))
tk.Button(root, text="📁 Browse Output", command=browse_output).pack(pady=5)
tk.Entry(root, textvariable=output_var, width=70).pack(padx=10, pady=5)

tk.Label(root, text="Select output formats:").pack(anchor="w", padx=10, pady=(10, 0))
for fmt, var in format_vars.items():
    tk.Checkbutton(root, text=fmt.upper(), variable=var).pack(anchor="w", padx=20)

tk.Checkbutton(root, text="🗃️ Merge PDFs into one", variable=merge_var).pack(anchor="w", padx=10, pady=10)
tk.Button(root, text="🚀 Start Conversion", command=threaded_conversion).pack(pady=20)

progress = ttk.Progressbar(root, length=500, mode='determinate')
progress.pack(pady=5)
progress_label = tk.Label(root, text="")
progress_label.pack(pady=5)

tk.Label(root, text="v2.1 | Created by Yoavs5771", font=("Arial", 8)).pack(side="bottom", pady=10)

# Enable drag & drop
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', handle_drop)

root.mainloop()
