import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from fpdf import FPDF
import pytesseract
import sys
import ctypes
import threading
from concurrent.futures import ThreadPoolExecutor
import queue
import logging
from multiprocessing import Value
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

class UTF8FPDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Header', 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, 'Page %s' % self.page_no(), 0, 0, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(4)

    def chapter_body(self, body):
        self.set_font('Arial', '', 12)
        self.multi_cell(0, 10, body)

class OCRApp:
    def __init__(self, master):
        self.master = master
        self.master.title("OCR App")
        self.master.geometry("800x600")

        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.language = tk.StringVar(value="pol")
        self.output_format = tk.StringVar(value="txt")
        self.progress_var = Value('d', 0.0)  # Zmiana na Value z multiprocessing
        self.progress_label_var = tk.StringVar()
        self.num_threads = tk.IntVar(value=4)
        self.ocr_finished = threading.Event()
        self.resume_flag = False

        self.set_tesseract_path()
        self.create_widgets()

        self.thread_data_queue = queue.Queue()
        self.lock = threading.Lock()
        self.pause_flag = False
        self.resume_event = threading.Event()

    def set_tesseract_path(self):
        expected_path = r"C:\Program Files (x86)\Tesseract-OCR"
        if not os.path.isdir(expected_path):
            messagebox.showwarning("Warning", f"The script is not running in the expected path ({expected_path}). Please ensure Tesseract is installed and the path is correctly set.")
        pytesseract.pytesseract.tesseract_cmd = os.path.join(expected_path, 'tesseract')

    def create_widgets(self):
    ttk.Label(self.master, text="Input Folder:").grid(column=0, row=0, sticky=tk.W)
    input_entry = ttk.Entry(self.master, textvariable=self.input_folder, width=50)
    input_entry.grid(column=1, row=0, sticky=(tk.W, tk.E))
    ttk.Button(self.master, text="Browse...", command=self.browse_input_folder).grid(column=2, row=0, sticky=tk.W)

    ttk.Label(self.master, text="Output Folder:").grid(column=0, row=1, sticky=tk.W)
    output_entry = ttk.Entry(self.master, textvariable=self.output_folder, width=50)
    output_entry.grid(column=1, row=1, sticky=(tk.W, tk.E))
    ttk.Button(self.master, text="Browse...", command=self.browse_output_folder).grid(column=2, row=1, sticky=tk.W)

    ttk.Label(self.master, text="Language:").grid(column=0, row=2, sticky=tk.W)
    ttk.Combobox(self.master, textvariable=self.language, values=["eng", "pol"], width=47).grid(column=1, row=2, sticky=tk.W)

    ttk.Label(self.master, text="Output Format:").grid(column=0, row=3, sticky=tk.W)
    ttk.Combobox(self.master, textvariable=self.output_format, values=["txt", "pdf"], width=47).grid(column=1, row=3, sticky=tk.W)

    ttk.Button(self.master, text="Start OCR", command=self.start_ocr).grid(column=0, row=4, sticky=tk.W, pady=4)
    ttk.Button(self.master, text="Pause", command=self.pause_ocr).grid(column=1, row=4, sticky=tk.W, pady=4)
    ttk.Button(self.master, text="Resume", command=self.resume_ocr).grid(column=2, row=4, sticky=tk.W, pady=4)

    progress = ttk.Progressbar(self.master, length=100, mode='determinate', variable=self.progress_var)
    progress.grid(column=0, row=5, columnspan=3, sticky=(tk.W, tk.E))
    ttk.Label(self.master, textvariable=self.progress_label_var).grid(column=1, row=6)

def browse_input_folder(self):
    folder_selected = filedialog.askdirectory()
    self.input_folder.set(folder_selected)

def browse_output_folder(self):
    folder_selected = filedialog.askdirectory()
    self.output_folder.set(folder_selected)

def start_ocr(self):
    if not self.input_folder.get() or not self.output_folder.get():
        messagebox.showwarning("Warning", "Please select both input and output folders.")
        return

    self.ocr_finished.clear()
    self.resume_flag = True
    self.resume_event.set()

    input_path = self.input_folder.get()
    output_path = self.output_folder.get()
    lang = self.language.get()
    format = self.output_format.get()

    images = [os.path.join(input_path, f) for f in os.listdir(input_path) if f.endswith(('.png', '.jpg', '.jpeg', '.tiff'))]

    if not images:
        messagebox.showinfo("Info", "No images found in the selected folder.")
        return

    total_images = len(images)
    self.progress_var.value = 0.0
    self.progress_label_var.set("0%")

    with ThreadPoolExecutor(max_workers=self.num_threads.get()) as executor:
        futures = [executor.submit(self.process_image, image, output_path, lang, format) for image in images]
        for future in futures:
            future.result()  # You can add more error handling here

    self.show_completed_message("OCR process has been completed.")

def process_image(self, image_path, output_path, lang, format):
    if self.pause_flag:
        self.resume_event.wait()  # Wait until the flag is set to False

    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img, lang=lang)

        base_filename = os.path.splitext(os.path.basename(image_path))[0]
        if format == "txt":
            with open(os.path.join(output_path, base_filename + ".txt"), "w", encoding="utf-8") as f:
                f.write(text)
        elif format == "pdf":
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
            pdf.set_font('DejaVu', '', 14)
            pdf.multi_cell(0, 10, text)
            pdf.output(os.path.join(output_path, base_filename + ".pdf"))

        with self.lock:
            self.progress_var.value += 100.0 / len(images)
            self.progress_label_var.set(f"{int(self.progress_var.value)}%")
            self.master.update_idletasks()

    except Exception as e:
        logging.error(f"Error processing file {image_path}: {e}")

def pause_ocr(self):
    self.pause_flag = True
    self.resume_event.clear()

def resume_ocr(self):
    self.pause_flag = False
    self.resume_event.set()

def show_completed_message(self, message):
    messagebox.showinfo("Completed", message)
