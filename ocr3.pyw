import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from fpdf import FPDF
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pytesseract
import sys
import ctypes
import threading
from functools import partial

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
        
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_rowconfigure(1, weight=1)
        self.master.grid_rowconfigure(2, weight=1)
        self.master.grid_rowconfigure(3, weight=1)
        self.master.grid_rowconfigure(4, weight=1)
        self.master.grid_rowconfigure(5, weight=1)
        self.master.grid_rowconfigure(6, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=1)
        self.master.grid_columnconfigure(2, weight=1)
        self.master.grid_columnconfigure(3, weight=1)

        ttk.Sizegrip(self.master).grid(row=7, column=3, sticky="se")

        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.language = tk.StringVar(value="pol")
        self.output_format = tk.StringVar(value="txt")
        self.progress_var = tk.DoubleVar()
        self.progress_label_var = tk.StringVar()
        self.num_threads = tk.StringVar(value="4")
        self.ocr_finished = False
        self.resume_flag = False

        self.set_tesseract_path()
        self.load_last_used_folders()
        self.create_widgets()

        self.thread_data = []
        self.threads = []
        self.lock = threading.Lock()
        self.pause_flag = False
        self.resume_event = threading.Event()

    def set_tesseract_path(self):
        expected_path = r"C:\Program Files (x86)\Tesseract-OCR"
        current_path = os.getcwd()

        if current_path != expected_path:
            messagebox.showwarning("Warning", f"The script is not running in the expected path ({expected_path}). Trying to move to this path.")

            try:
                os.chdir(expected_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to move to the expected path: {e}")
                sys.exit()

    def create_widgets(self):
        tk.Label(self.master, text="Folder wejściowy:").grid(row=0, column=0, sticky="e")
        tk.Label(self.master, text="Folder wyjściowy:").grid(row=1, column=0, sticky="e")
        tk.Label(self.master, text="Język Tesseracta:").grid(row=2, column=0, sticky="e")
        tk.Label(self.master, text="Format wyjściowy:").grid(row=3, column=0, sticky="e")
        tk.Label(self.master, text="Liczba wątków:").grid(row=4, column=0, sticky="e")
        tk.Label(self.master, text="Postęp:").grid(row=6, column=0, sticky="e")

        tk.Entry(self.master, textvariable=self.input_folder, width=40, state="readonly").grid(row=0, column=1, padx=5, pady=5, sticky="we")
        tk.Entry(self.master, textvariable=self.output_folder, width=40, state="readonly").grid(row=1, column=1, padx=5, pady=5, sticky="we")

        tk.Button(self.master, text="Wybierz folder wejściowy", command=self.browse_input_folder).grid(row=0, column=2, padx=5, pady=10, sticky="we")
        tk.Button(self.master, text="Wybierz folder wyjściowy", command=self.browse_output_folder).grid(row=1, column=2, padx=5, pady=10, sticky="we")

        languages = ["eng", "pol"]
        tk.OptionMenu(self.master, self.language, *languages).grid(row=2, column=1, padx=5, pady=10, sticky="we")

        formats = ["txt", "pdf"]
        tk.OptionMenu(self.master, self.output_format, *formats).grid(row=3, column=1, padx=5, pady=10, sticky="we")

        tk.OptionMenu(self.master, self.num_threads, *map(str, range(1, 31))).grid(row=4, column=1, padx=5, pady=10, sticky="we")

        tk.Button(self.master, text="Uruchom OCR", command=self.start_ocr).grid(row=5, column=1, pady=10, sticky="we")
        tk.Button(self.master, text="Wstrzymaj OCR", command=self.pause_ocr).grid(row=5, column=2, pady=10, sticky="we")
        tk.Button(self.master, text="Wznów OCR", command=self.resume_ocr).grid(row=5, column=3, pady=10, sticky="we")

        ttk.Progressbar(self.master, variable=self.progress_var, length=200, mode='determinate').grid(row=6, column=1, pady=10, sticky="we")
        tk.Label(self.master, textvariable=self.progress_label_var).grid(row=6, column=2, pady=10, sticky="we")

    def browse_input_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.input_folder.set(folder_selected)
            self.save_last_used_folders()

    def browse_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder.set(folder_selected)
            self.save_last_used_folders()

    def update_progress(self, current, total):
        progress_percentage = (current / total) * 100
        self.progress_var.set(progress_percentage)
        self.progress_label_var.set(f"{int(progress_percentage)}%")
        self.master.update_idletasks()

    def start_ocr(self):
        if self.resume_flag:
            self.resume_flag = False
            self.resume_event.set()
            messagebox.showinfo("Info", "Wznowano OCR.")
            return

        input_folder = self.input_folder.get()
        output_folder = self.output_folder.get()
        language = self.language.get()
        output_format = self.output_format.get()
        num_threads = int(self.num_threads.get())

        if not input_folder or not output_folder:
            messagebox.showerror("Błąd", "Wybierz folder wejściowy i wyjściowy.")
            return

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        image_files = [filename for filename in os.listdir(input_folder) if filename.lower().endswith((".jpg", ".jpeg", ".png"))]
        total_images = len(image_files)

        self.thread_data = [(i, filename, input_folder, output_folder, language, output_format, total_images) for i, filename in enumerate(image_files, start=1)]

        self.update_progress(0, total_images)

        # Utwórz wątki
        self.threads = [threading.Thread(target=self.process_next_image) for _ in range(num_threads)]

        # Uruchom wątki
        self.resume_event.set()
        for thread in self.threads:
            thread.start()

    def process_next_image(self):
        while self.thread_data and self.resume_event.is_set():
            thread_args = self.thread_data.pop(0)
            self.process_image(thread_args)

        # Sprawdź, czy wszystkie wątki zostały zakończone
        if not any(thread.is_alive() for thread in self.threads) and not self.ocr_finished:
            if self.progress_var.get() == 100:  # Dodana warunek sprawdzający, czy postęp wynosi 100%
                self.show_completed_message("Proces OCR został zakończony.")
                self.ocr_finished = True

    def process_image(self, thread_args):
        i, filename, input_folder, output_folder, language, output_format, total_images = thread_args

        input_path = os.path.join(input_folder, filename)
        image = Image.open(input_path)
        text = pytesseract.image_to_string(image, lang=language)

        if output_format == "txt":
            output_filename = os.path.splitext(filename)[0] + ".txt"
            output_path = os.path.join(output_folder, output_filename)

            with open(output_path, 'w', encoding='utf-8') as txt_file:
                txt_file.write(text)
        elif output_format == "pdf":
            output_filename = os.path.splitext(filename)[0] + ".pdf"
            output_path = os.path.join(output_folder, output_filename)

            self.save_text_to_pdf(text, output_path)

        with self.lock:
            self.update_progress(i, total_images)

    def save_text_to_pdf(self, text, output_path):
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, text)
        pdf.output(output_path)

    def load_last_used_folders(self):
        config_file = 'last_used_folders.ini'
        if os.path.exists(config_file):
            with open(config_file, 'r') as configfile:
                lines = configfile.readlines()
                if len(lines) == 2:
                    self.input_folder.set(lines[0].strip())
                    self.output_folder.set(lines[1].strip())

    def save_last_used_folders(self):
        config_file = 'last_used_folders.ini'
        with open(config_file, 'w') as configfile:
            configfile.write(f"{self.input_folder.get()}\n{self.output_folder.get()}")

    def pause_ocr(self):
        self.pause_flag = True
        self.resume_event.clear()
        messagebox.showinfo("Info", "OCR zostało wstrzymane. Wznow OCR, aby kontynuować.")

    def resume_ocr(self):
        if self.pause_flag:
            self.pause_flag = False
            self.resume_event.set()
            messagebox.showinfo("Info", "Wznowano OCR.")

    def show_completed_message(self, message):
        messagebox.showinfo("Zakończono", message)

if __name__ == "__main__":
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()
