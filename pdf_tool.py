import argparse
import PyPDF2
import os
from pdf2image import convert_from_path
from pdf2docx import Converter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from pikepdf import Pdf, Encryption
import io
import subprocess
import pdfkit

def merge_pdfs(input_files, output_file):
    pdf_merger = PyPDF2.PdfMerger()
    for pdf in input_files:
        pdf_merger.append(pdf)
    pdf_merger.write(output_file)
    pdf_merger.close()
    print(f"PDF files {input_files} merged into {output_file}")

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
        print(f"Pages {start} to {end} from {input_file} saved as {output_filename}")
    elif pages_per_split:
        pages_per_split = int(pages_per_split)
        for i in range(0, len(pdf_reader.pages), pages_per_split):
            pdf_writer = PyPDF2.PdfWriter()
            for j in range(i, min(i + pages_per_split, len(pdf_reader.pages))):
                pdf_writer.add_page(pdf_reader.pages[j])
            output_filename = os.path.join(output_dir, f'pages_{i + 1}_to_{min(i + pages_per_split, len(pdf_reader.pages))}.pdf')
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
            print(f"Pages {i + 1} to {min(i + pages_per_split, len(pdf_reader.pages))} from {input_file} saved as {output_filename}")
    else:
        for page_num in range(len(pdf_reader.pages)):
            pdf_writer = PyPDF2.PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page_num])
            output_filename = os.path.join(output_dir, f'page_{page_num + 1}.pdf')
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
            print(f"Page {page_num + 1} from {input_file} saved as {output_filename}")

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

    print(f"Original size: {original_size} bytes")
    print(f"Target size: {target_size} bytes")
    print(f"Compression resulted in {current_size} bytes.")

    if current_size > target_size:
        print(f"Unable to compress {input_file} to the target size {target_size} bytes. Current size is {current_size} bytes.")
    else:
        print(f"PDF file {input_file} compressed to {output_file} with size {current_size} bytes, meeting the target size {target_size} bytes.")

def pdf_to_images(input_file, output_dir):
    """Convert PDF pages to images."""
    images = convert_from_path(input_file)
    for i, image in enumerate(images):
        image_filename = os.path.join(output_dir, f'page_{i + 1}.jpg')
        image.save(image_filename, 'JPEG')
        print(f"Page {i + 1} from {input_file} saved as {image_filename}")

def pdf_to_word(input_file, output_file):
    """Convert PDF to Word document."""
    cv = Converter(input_file)
    cv.convert(output_file)
    cv.close()
    print(f"PDF file {input_file} converted to Word document {output_file}")

def rotate_pdf(input_file, output_file, angle):
    """Rotate pages of a PDF file."""
    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer = PyPDF2.PdfWriter()
    for page in pdf_reader.pages:
        page.rotate_clockwise(angle)
        pdf_writer.add_page(page)
    with open(output_file, 'wb') as out:
        pdf_writer.write(out)
    print(f"PDF file {input_file} rotated by {angle} degrees and saved as {output_file}")

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
    print(f"Page numbers added to {input_file} and saved as {output_file}")

def repair_pdf(input_file, output_file):
    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer = PyPDF2.PdfWriter()

    for page in pdf_reader.pages:
        pdf_writer.add_page(page)

    with open(output_file, 'wb') as out:
        pdf_writer.write(out)
    print(f"PDF file {input_file} repaired and saved as {output_file}")

def unlock_pdf(input_file, output_file, password):
    pdf_reader = PyPDF2.PdfReader(input_file)
    if pdf_reader.is_encrypted:
        pdf_reader.decrypt(password)

    pdf_writer = PyPDF2.PdfWriter()
    for page in pdf_reader.pages:
        pdf_writer.add_page(page)

    with open(output_file, 'wb') as out:
        pdf_writer.write(out)
    print(f"Password protection removed from {input_file} and saved as {output_file}")

def protect_pdf(input_file, output_file, password):
    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer = PyPDF2.PdfWriter()

    for page in pdf_reader.pages:
        pdf_writer.add_page(page)

    pdf_writer.encrypt(password)
    with open(output_file, 'wb') as out:
        pdf_writer.write(out)
    print(f"PDF file {input_file} protected with a password and saved as {output_file}")

def url_to_pdf(url, output_file):
    pdfkit.from_url(url, output_file)
    print(f"URL {url} converted to PDF and saved as {output_file}")

def main():
    parser = argparse.ArgumentParser(
        description="PDF manipulation tool",
        epilog="Example usage: python pdf_tool.py merge file1.pdf file2.pdf output.pdf"
    )
    subparsers = parser.add_subparsers(title="Commands", dest="command")

    # Merge command
    merge_parser = subparsers.add_parser("merge", help="Merge multiple PDF files")
    merge_parser.add_argument("input_files", nargs='+', help="Input PDF files to merge")
    merge_parser.add_argument("output_file", help="Output merged PDF file")

    # Split command
    split_parser = subparsers.add_parser("split", help="Split a PDF file into individual pages or specified pages per file or range of pages")
    split_parser.add_argument("input_file", help="Input PDF file to split")
    split_parser.add_argument("output_dir", help="Output directory for split PDF files")
    split_parser.add_argument("--pages_per_split", type=int, help="Number of pages per split file")
    split_parser.add_argument("--page_range", type=str, help="Range of pages to extract (e.g., 1-3)")

    # Compress command
    compress_parser = subparsers.add_parser("compress", help="Compress a PDF file to a specified percentage of its original size")
    compress_parser.add_argument("input_file", help="Input PDF file to compress")
    compress_parser.add_argument("output_file", help="Output compressed PDF file")
    compress_parser.add_argument("compression_percentage", type=int, help="Compression percentage (e.g., 60 for 60% of the original size)")

    # PDF to Images command
    pdf_to_images_parser = subparsers.add_parser("pdf_to_images", help="Convert PDF pages to images")
    pdf_to_images_parser.add_argument("input_file", help="Input PDF file to convert to images")
    pdf_to_images_parser.add_argument("output_dir", help="Output directory for images")

    # PDF to Word command
    pdf_to_word_parser = subparsers.add_parser("pdf_to_word", help="Convert PDF to Word document")
    pdf_to_word_parser.add_argument("input_file", help="Input PDF file to convert to Word")
    pdf_to_word_parser.add_argument("output_file", help="Output Word document file")

    # Rotate command
    rotate_parser = subparsers.add_parser("rotate", help="Rotate pages of a PDF file")
    rotate_parser.add_argument("input_file", help="Input PDF file to rotate")
    rotate_parser.add_argument("output_file", help="Output rotated PDF file")
    rotate_parser.add_argument("angle", type=int, help="Angle to rotate PDF pages (clockwise)")

    # Add page numbers command
    add_page_numbers_parser = subparsers.add_parser("add_page_numbers", help="Add page numbers to a PDF file")
    add_page_numbers_parser.add_argument("input_file", help="Input PDF file")
    add_page_numbers_parser.add_argument("output_file", help="Output PDF file with page numbers")

    # Repair command
    repair_parser = subparsers.add_parser("repair", help="Repair a PDF file")
    repair_parser.add_argument("input_file", help="Input PDF file to repair")
    repair_parser.add_argument("output_file", help="Output repaired PDF file")

    # Unlock command
    unlock_parser = subparsers.add_parser("unlock", help="Unlock a password-protected PDF file")
    unlock_parser.add_argument("input_file", help="Input PDF file to unlock")
    unlock_parser.add_argument("output_file", help="Output unlocked PDF file")
    unlock_parser.add_argument("password", help="Password for the PDF file")

    # Protect command
    protect_parser = subparsers.add_parser("protect", help="Protect a PDF file with a password")
    protect_parser.add_argument("input_file", help="Input PDF file to protect")
    protect_parser.add_argument("output_file", help="Output protected PDF file")
    protect_parser.add_argument("password", help="Password to protect the PDF file")

    # URL to PDF command
    url_to_pdf_parser = subparsers.add_parser("url_to_pdf", help="Convert a webpage to PDF")
    url_to_pdf_parser.add_argument("url", help="URL of the webpage to convert to PDF")
    url_to_pdf_parser.add_argument("output_file", help="Output PDF file")

    args = parser.parse_args()

    if args.command == "merge":
        merge_pdfs(args.input_files, args.output_file)
    elif args.command == "split":
        split_pdf(args.input_file, args.output_dir, args.pages_per_split, args.page_range)
    elif args.command == "compress":
        compress_pdf(args.input_file, args.output_file, args.compression_percentage)
    elif args.command == "pdf_to_images":
        pdf_to_images(args.input_file, args.output_dir)
    elif args.command == "pdf_to_word":
        pdf_to_word(args.input_file, args.output_file)
    elif args.command == "rotate":
        rotate_pdf(args.input_file, args.output_file, args.angle)
    elif args.command == "add_page_numbers":
        add_page_numbers(args.input_file, args.output_file)
    elif args.command == "repair":
        repair_pdf(args.input_file, args.output_file)
    elif args.command == "unlock":
        unlock_pdf(args.input_file, args.output_file, args.password)
    elif args.command == "protect":
        protect_pdf(args.input_file, args.output_file, args.password)
    elif args.command == "url_to_pdf":
        url_to_pdf(args.url, args.output_file)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
