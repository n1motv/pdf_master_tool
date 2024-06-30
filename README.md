# PDF Master Tool

PDF Master Tool is a comprehensive PDF manipulation tool written in Python. It allows users to merge, split, compress, convert, rotate, add page numbers, repair, lock, and unlock PDF files effortlessly.

## Features
- Merge multiple PDF files into one.
- Split a PDF file into individual pages or custom page ranges.
- Compress PDF files to reduce their size.
- Convert PDF pages to images.
- Convert PDFs to Word documents.
- Rotate PDF pages.
- Add page numbers to PDF files.
- Repair corrupted PDF files.
- Unlock password-protected PDFs.
- Protect PDF files with a password.
- Convert URL to PDF file

## Installation
1. Clone the repository:
    ```sh
    git clone https://github.com/n1motv/pdf_master_tool
    ```
2. Navigate to the project directory:
    ```sh
    cd pdf_master_tool
    ```
3. Install the required dependencies:
   ```sh
    sudo apt-get update
    sudo apt-get install -y poppler-utils ghostscript
    sudo apt-get install wkhtmltopdf
    ```
    ```sh
    pip install -r requirements.txt
    ```

## Usage
```sh
# Merging PDFs
python pdf_tool.py merge file1.pdf file2.pdf output.pdf

# Splitting PDFs
# Split by pages per file:
python pdf_tool.py split input.pdf output_dir --pages_per_split 5

# Split by page range:
python pdf_tool.py split input.pdf output_dir --page_range 1-3

# Compressing PDFs
python pdf_tool.py compress input.pdf output.pdf 60

# Converting PDFs to Images
python pdf_tool.py pdf_to_images input.pdf output_dir

# Converting PDFs to Word
python pdf_tool.py pdf_to_word input.pdf output.docx

# Rotating PDFs
python pdf_tool.py rotate input.pdf output.pdf 90

# Adding Page Numbers
python pdf_tool.py add_page_numbers input.pdf output.pdf

# Repairing PDFs
python pdf_tool.py repair input.pdf output.pdf

# Unlocking PDFs
python pdf_tool.py unlock input.pdf output.pdf password

# Protecting PDFs
python pdf_tool.py protect input.pdf output.pdf password

#Converting URL to PDF
python pdf_tool.py url_to_pdf www.example.com output.pdf
```

## Contributing
Contributions are welcome! Please fork the repository and submit a pull request.
