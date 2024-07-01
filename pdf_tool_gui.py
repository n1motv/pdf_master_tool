import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ttkthemes import ThemedTk
import os
import subprocess
from pdf2image import convert_from_path
from pdf2docx import Converter
import PyPDF2
import pdfkit
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4


# Function definitions as provided previously

def merge_pdfs(input_files, output_file):
    pdf_merger = PyPDF2.PdfMerger()
    for pdf in input_files:
        pdf_merger.append(pdf)
    pdf_merger.write(output_file)
    pdf_merger.close()
    messagebox.showinfo("Success", f"PDF files merged into {output_file}")


def split_pdf(input_file, output_dir, pages_per_split=None, page_range=None):
    pdf_reader = PyPDF2.PdfReader(input_file)
    if page_range:
        start, end = map(int, page_range.split('-'))
        pdf_writer = PyPDF2.PdfWriter()
        for page_num in range(start - 1, end):
            pdf_writer.add_page(pdf_reader.pages[page_num])
        output_filename = os.path.join(output_dir, f'pages_{start}_to_{end}.pdf')
        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)
        messagebox.showinfo("Success", f"Pages {start} to {end} saved as {output_filename}")
    elif pages_per_split:
        pages_per_split = int(pages_per_split)
        for i in range(0, len(pdf_reader.pages), pages_per_split):
            pdf_writer = PyPDF2.PdfWriter()
            for j in range(i, min(i + pages_per_split, len(pdf_reader.pages))):
                pdf_writer.add_page(pdf_reader.pages[j])
            output_filename = os.path.join(output_dir,
                                           f'pages_{i + 1}_to_{min(i + pages_per_split, len(pdf_reader.pages))}.pdf')
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
            messagebox.showinfo("Success",
                                f"Pages {i + 1} to {min(i + pages_per_split, len(pdf_reader.pages))} saved as {output_filename}")
    else:
        for page_num in range(len(pdf_reader.pages)):
            pdf_writer = PyPDF2.PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page_num])
            output_filename = os.path.join(output_dir, f'page_{page_num + 1}.pdf')
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
            messagebox.showinfo("Success", f"Page {page_num + 1} saved as {output_filename}")


def compress_pdf(input_file, output_file, compression_percentage):
    original_size = os.path.getsize(input_file)
    target_size = int(original_size * (compression_percentage / 100))
    gs_command = [
        'gs',
        '-sDEVICE=pdfwrite',
        '-dCompatibilityLevel=1.4',
        '-dPDFSETTINGS=/ebook',
        '-dNOPAUSE',
        '-dQUIET',
        '-dBATCH',
        f'-sOutputFile={output_file}',
        input_file
    ]
    subprocess.run(gs_command, check=True)
    current_size = os.path.getsize(output_file)
    if current_size > target_size:
        messagebox.showwarning("Warning",
                               f"Unable to compress to target size {target_size} bytes. Current size is {current_size} bytes.")
    else:
        messagebox.showinfo("Success", f"Compressed to {current_size} bytes.")


def pdf_to_images(input_file, output_dir):
    images = convert_from_path(input_file)
    for i, image in enumerate(images):
        image_filename = os.path.join(output_dir, f'page_{i + 1}.jpg')
        image.save(image_filename, 'JPEG')
        messagebox.showinfo("Success", f"Page {i + 1} saved as {image_filename}")


def pdf_to_word(input_file, output_file):
    cv = Converter(input_file)
    cv.convert(output_file)
    cv.close()
    messagebox.showinfo("Success", f"Converted to Word document {output_file}")


def rotate_pdf(input_file, output_file, angle):
    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer = PyPDF2.PdfWriter()
    for page in pdf_reader.pages:
        page.rotate_clockwise(angle)
        pdf_writer.add_page(page)
    with open(output_file, 'wb') as out:
        pdf_writer.write(out)
    messagebox.showinfo("Success", f"Rotated by {angle} degrees and saved as {output_file}")


def add_page_numbers(input_file, output_file):
    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer = PyPDF2.PdfWriter()
    for i, page in enumerate(pdf_reader.pages):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        can.drawString(500, 10, str(i + 1))
        can.save()
        packet.seek(0)
        watermark = PyPDF2.PdfReader(packet).pages[0]
        page.merge_page(watermark)
        pdf_writer.add_page(page)
    with open(output_file, 'wb') as out:
        pdf_writer.write(out)
    messagebox.showinfo("Success", f"Page numbers added and saved as {output_file}")


def repair_pdf(input_file, output_file):
    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer = PyPDF2.PdfWriter()
    for page in pdf_reader.pages:
        pdf_writer.add_page(page)
    with open(output_file, 'wb') as out:
        pdf_writer.write(out)
    messagebox.showinfo("Success", f"Repaired and saved as {output_file}")


def unlock_pdf(input_file, output_file, password):
    pdf_reader = PyPDF2.PdfReader(input_file)
    if pdf_reader.is_encrypted:
        pdf_reader.decrypt(password)
    pdf_writer = PyPDF2.PdfWriter()
    for page in pdf_reader.pages:
        pdf_writer.add_page(page)
    with open(output_file, 'wb') as out:
        pdf_writer.write(out)
    messagebox.showinfo("Success", f"Password protection removed and saved as {output_file}")


def protect_pdf(input_file, output_file, password):
    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer = PyPDF2.PdfWriter()
    for page in pdf_reader.pages:
        pdf_writer.add_page(page)
    pdf_writer.encrypt(password)
    with open(output_file, 'wb') as out:
        pdf_writer.write(out)
    messagebox.showinfo("Success", f"Protected with password and saved as {output_file}")


def url_to_pdf(url, output_file):
    config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
    pdfkit.from_url(url, output_file, configuration=config)
    messagebox.showinfo("Success", f"URL converted to PDF and saved as {output_file}")


# GUI Application

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Manipulation Tool")
        self.create_styles()
        self.create_widgets()

    def create_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        # Define colors
        primary_color = "#34495E"
        secondary_color = "#2C3E50"
        accent_color = "#1ABC9C"
        text_color = "#ECF0F1"
        entry_bg_color = "#2C3E50"
        entry_fg_color = "#ECF0F1"

        style.configure('TFrame', background=secondary_color)
        style.configure('TLabel', background=secondary_color, foreground=text_color, font=('Arial', 12))
        style.configure('TButton', background=primary_color, foreground=text_color, font=('Arial', 12))
        style.configure('TEntry', fieldbackground=entry_bg_color, foreground=entry_fg_color)

        style.map('TButton', background=[('active', accent_color)])

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=10, expand=True)

        self.create_merge_tab()
        self.create_split_tab()
        self.create_compress_tab()
        self.create_convert_tab()
        self.create_rotate_tab()
        self.create_add_page_numbers_tab()
        self.create_repair_tab()
        self.create_unlock_tab()
        self.create_protect_tab()
        self.create_url_to_pdf_tab()

    def create_merge_tab(self):
        self.merge_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.merge_tab, text="Merge PDFs")

        ttk.Label(self.merge_tab, text="Select PDF files to merge:").pack(pady=5)
        self.merge_files = []
        self.merge_files_listbox = tk.Listbox(self.merge_tab, selectmode=tk.MULTIPLE, background='#34495E',
                                              foreground='#ECF0F1')
        self.merge_files_listbox.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        ttk.Button(self.merge_tab, text="Add Files", command=self.add_merge_files).pack(pady=5)
        ttk.Button(self.merge_tab, text="Remove Selected", command=self.remove_merge_files).pack(pady=5)
        ttk.Button(self.merge_tab, text="Merge", command=self.merge_files_action).pack(pady=5)

    def create_split_tab(self):
        self.split_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.split_tab, text="Split PDF")

        ttk.Label(self.split_tab, text="Select PDF file to split:").pack(pady=5)
        self.split_file_entry = ttk.Entry(self.split_tab)
        self.split_file_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.split_tab, text="Browse", command=self.browse_split_file).pack(pady=5)

        ttk.Label(self.split_tab, text="Output directory:").pack(pady=5)
        self.split_output_dir_entry = ttk.Entry(self.split_tab)
        self.split_output_dir_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.split_tab, text="Browse", command=self.browse_split_output_dir).pack(pady=5)

        self.split_pages_per_split = ttk.Entry(self.split_tab)
        self.split_pages_per_split.pack(pady=5, padx=10, fill=tk.X)
        ttk.Label(self.split_tab, text="Pages per split (leave blank for single pages)").pack(pady=5)

        self.split_page_range = ttk.Entry(self.split_tab)
        self.split_page_range.pack(pady=5, padx=10, fill=tk.X)
        ttk.Label(self.split_tab, text="Page range (e.g., 1-3)").pack(pady=5)

        ttk.Button(self.split_tab, text="Split", command=self.split_file_action).pack(pady=5)

    def create_compress_tab(self):
        self.compress_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.compress_tab, text="Compress PDF")

        ttk.Label(self.compress_tab, text="Select PDF file to compress:").pack(pady=5)
        self.compress_file_entry = ttk.Entry(self.compress_tab)
        self.compress_file_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.compress_tab, text="Browse", command=self.browse_compress_file).pack(pady=5)

        ttk.Label(self.compress_tab, text="Output file:").pack(pady=5)
        self.compress_output_file_entry = ttk.Entry(self.compress_tab)
        self.compress_output_file_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Label(self.compress_tab, text="Compression percentage:").pack(pady=5)
        self.compress_percentage_entry = ttk.Entry(self.compress_tab)
        self.compress_percentage_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Button(self.compress_tab, text="Compress", command=self.compress_file_action).pack(pady=5)

    def create_convert_tab(self):
        self.convert_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.convert_tab, text="Convert PDF")

        ttk.Label(self.convert_tab, text="Select PDF file to convert:").pack(pady=5)
        self.convert_file_entry = ttk.Entry(self.convert_tab)
        self.convert_file_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.convert_tab, text="Browse", command=self.browse_convert_file).pack(pady=5)

        self.convert_to_images_tab()
        self.convert_to_word_tab()

    def convert_to_images_tab(self):
        self.convert_images_tab = ttk.Frame(self.convert_tab)
        self.convert_images_tab.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        ttk.Label(self.convert_images_tab, text="Convert to Images").pack(pady=5)

        ttk.Label(self.convert_images_tab, text="Output directory:").pack(pady=5)
        self.convert_output_images_dir_entry = ttk.Entry(self.convert_images_tab)
        self.convert_output_images_dir_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.convert_images_tab, text="Browse", command=self.browse_convert_output_images_dir).pack(pady=5)
        ttk.Button(self.convert_images_tab, text="Convert to Images", command=self.convert_to_images_action).pack(
            pady=5)

    def convert_to_word_tab(self):
        self.convert_word_tab = ttk.Frame(self.convert_tab)
        self.convert_word_tab.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        ttk.Label(self.convert_word_tab, text="Convert to Word").pack(pady=5)

        ttk.Label(self.convert_word_tab, text="Output file:").pack(pady=5)
        self.convert_output_word_file_entry = ttk.Entry(self.convert_word_tab)
        self.convert_output_word_file_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.convert_word_tab, text="Convert to Word", command=self.convert_to_word_action).pack(pady=5)

    def create_rotate_tab(self):
        self.rotate_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.rotate_tab, text="Rotate PDF")

        ttk.Label(self.rotate_tab, text="Select PDF file to rotate:").pack(pady=5)
        self.rotate_file_entry = ttk.Entry(self.rotate_tab)
        self.rotate_file_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.rotate_tab, text="Browse", command=self.browse_rotate_file).pack(pady=5)

        ttk.Label(self.rotate_tab, text="Output file:").pack(pady=5)
        self.rotate_output_file_entry = ttk.Entry(self.rotate_tab)
        self.rotate_output_file_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Label(self.rotate_tab, text="Rotation angle:").pack(pady=5)
        self.rotate_angle_entry = ttk.Entry(self.rotate_tab)
        self.rotate_angle_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Button(self.rotate_tab, text="Rotate", command=self.rotate_file_action).pack(pady=5)

    def create_add_page_numbers_tab(self):
        self.add_page_numbers_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.add_page_numbers_tab, text="Add Page Numbers")

        ttk.Label(self.add_page_numbers_tab, text="Select PDF file to add page numbers:").pack(pady=5)
        self.add_page_numbers_file_entry = ttk.Entry(self.add_page_numbers_tab)
        self.add_page_numbers_file_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.add_page_numbers_tab, text="Browse", command=self.browse_add_page_numbers_file).pack(pady=5)

        ttk.Label(self.add_page_numbers_tab, text="Output file:").pack(pady=5)
        self.add_page_numbers_output_file_entry = ttk.Entry(self.add_page_numbers_tab)
        self.add_page_numbers_output_file_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Button(self.add_page_numbers_tab, text="Add Page Numbers", command=self.add_page_numbers_action).pack(
            pady=5)

    def create_repair_tab(self):
        self.repair_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.repair_tab, text="Repair PDF")

        ttk.Label(self.repair_tab, text="Select PDF file to repair:").pack(pady=5)
        self.repair_file_entry = ttk.Entry(self.repair_tab)
        self.repair_file_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.repair_tab, text="Browse", command=self.browse_repair_file).pack(pady=5)

        ttk.Label(self.repair_tab, text="Output file:").pack(pady=5)
        self.repair_output_file_entry = ttk.Entry(self.repair_tab)
        self.repair_output_file_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Button(self.repair_tab, text="Repair", command=self.repair_file_action).pack(pady=5)

    def create_unlock_tab(self):
        self.unlock_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.unlock_tab, text="Unlock PDF")

        ttk.Label(self.unlock_tab, text="Select PDF file to unlock:").pack(pady=5)
        self.unlock_file_entry = ttk.Entry(self.unlock_tab)
        self.unlock_file_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.unlock_tab, text="Browse", command=self.browse_unlock_file).pack(pady=5)

        ttk.Label(self.unlock_tab, text="Output file:").pack(pady=5)
        self.unlock_output_file_entry = ttk.Entry(self.unlock_tab)
        self.unlock_output_file_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Label(self.unlock_tab, text="Password:").pack(pady=5)
        self.unlock_password_entry = ttk.Entry(self.unlock_tab, show='*')
        self.unlock_password_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Button(self.unlock_tab, text="Unlock", command=self.unlock_file_action).pack(pady=5)

    def create_protect_tab(self):
        self.protect_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.protect_tab, text="Protect PDF")

        ttk.Label(self.protect_tab, text="Select PDF file to protect:").pack(pady=5)
        self.protect_file_entry = ttk.Entry(self.protect_tab)
        self.protect_file_entry.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.protect_tab, text="Browse", command=self.browse_protect_file).pack(pady=5)

        ttk.Label(self.protect_tab, text="Output file:").pack(pady=5)
        self.protect_output_file_entry = ttk.Entry(self.protect_tab)
        self.protect_output_file_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Label(self.protect_tab, text="Password:").pack(pady=5)
        self.protect_password_entry = ttk.Entry(self.protect_tab, show='*')
        self.protect_password_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Button(self.protect_tab, text="Protect", command=self.protect_file_action).pack(pady=5)

    def create_url_to_pdf_tab(self):
        self.url_to_pdf_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.url_to_pdf_tab, text="URL to PDF")

        ttk.Label(self.url_to_pdf_tab, text="Enter URL:").pack(pady=5)
        self.url_entry = ttk.Entry(self.url_to_pdf_tab)
        self.url_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Label(self.url_to_pdf_tab, text="Output file:").pack(pady=5)
        self.url_output_file_entry = ttk.Entry(self.url_to_pdf_tab)
        self.url_output_file_entry.pack(pady=5, padx=10, fill=tk.X)

        ttk.Button(self.url_to_pdf_tab, text="Convert", command=self.url_to_pdf_action).pack(pady=5)

    def add_merge_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for file in files:
            if file not in self.merge_files:
                self.merge_files.append(file)
                self.merge_files_listbox.insert(tk.END, file)

    def remove_merge_files(self):
        selected_files = self.merge_files_listbox.curselection()
        for index in selected_files[::-1]:
            self.merge_files_listbox.delete(index)
            del self.merge_files[index]

    def merge_files_action(self):
        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_file:
            merge_pdfs(self.merge_files, output_file)

    def browse_split_file(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.split_file_entry.delete(0, tk.END)
            self.split_file_entry.insert(0, file)

    def browse_split_output_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.split_output_dir_entry.delete(0, tk.END)
            self.split_output_dir_entry.insert(0, dir_path)

    def split_file_action(self):
        input_file = self.split_file_entry.get()
        output_dir = self.split_output_dir_entry.get()
        pages_per_split = self.split_pages_per_split.get()
        page_range = self.split_page_range.get()
        if input_file and output_dir:
            split_pdf(input_file, output_dir, pages_per_split, page_range)

    def browse_compress_file(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.compress_file_entry.delete(0, tk.END)
            self.compress_file_entry.insert(0, file)

    def compress_file_action(self):
        input_file = self.compress_file_entry.get()
        output_file = self.compress_output_file_entry.get()
        compression_percentage = int(self.compress_percentage_entry.get())
        if input_file and output_file and compression_percentage:
            compress_pdf(input_file, output_file, compression_percentage)

    def browse_convert_file(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.convert_file_entry.delete(0, tk.END)
            self.convert_file_entry.insert(0, file)

    def browse_convert_output_images_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.convert_output_images_dir_entry.delete(0, tk.END)
            self.convert_output_images_dir_entry.insert(0, dir_path)

    def convert_to_images_action(self):
        input_file = self.convert_file_entry.get()
        output_dir = self.convert_output_images_dir_entry.get()
        if input_file and output_dir:
            pdf_to_images(input_file, output_dir)

    def convert_to_word_action(self):
        input_file = self.convert_file_entry.get()
        output_file = self.convert_output_word_file_entry.get()
        if input_file and output_file:
            pdf_to_word(input_file, output_file)

    def browse_rotate_file(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.rotate_file_entry.delete(0, tk.END)
            self.rotate_file_entry.insert(0, file)

    def rotate_file_action(self):
        input_file = self.rotate_file_entry.get()
        output_file = self.rotate_output_file_entry.get()
        angle = int(self.rotate_angle_entry.get())
        if input_file and output_file and angle:
            rotate_pdf(input_file, output_file, angle)

    def browse_add_page_numbers_file(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.add_page_numbers_file_entry.delete(0, tk.END)
            self.add_page_numbers_file_entry.insert(0, file)

    def add_page_numbers_action(self):
        input_file = self.add_page_numbers_file_entry.get()
        output_file = self.add_page_numbers_output_file_entry.get()
        if input_file and output_file:
            add_page_numbers(input_file, output_file)

    def browse_repair_file(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.repair_file_entry.delete(0, tk.END)
            self.repair_file_entry.insert(0, file)

    def repair_file_action(self):
        input_file = self.repair_file_entry.get()
        output_file = self.repair_output_file_entry.get()
        if input_file and output_file:
            repair_pdf(input_file, output_file)

    def browse_unlock_file(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.unlock_file_entry.delete(0, tk.END)
            self.unlock_file_entry.insert(0, file)

    def unlock_file_action(self):
        input_file = self.unlock_file_entry.get()
        output_file = self.unlock_output_file_entry.get()
        password = self.unlock_password_entry.get()
        if input_file and output_file and password:
            unlock_pdf(input_file, output_file, password)

    def browse_protect_file(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.protect_file_entry.delete(0, tk.END)
            self.protect_file_entry.insert(0, file)

    def protect_file_action(self):
        input_file = self.protect_file_entry.get()
        output_file = self.protect_output_file_entry.get()
        password = self.protect_password_entry.get()
        if input_file and output_file and password:
            protect_pdf(input_file, output_file, password)

    def url_to_pdf_action(self):
        url = self.url_entry.get()
        output_file = self.url_output_file_entry.get()
        if url and output_file:
            url_to_pdf(url, output_file)


if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    app = PDFToolApp(root)
    root.mainloop()
