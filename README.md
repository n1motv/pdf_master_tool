PDFMasterTool
PDFMasterTool is a comprehensive PDF manipulation tool written in Python. It allows users to merge, split, compress, convert, rotate, add page numbers, repair, lock, and unlock PDF files effortlessly.

Features
Merge multiple PDF files into one.
Split a PDF file into individual pages or custom page ranges.
Compress PDF files to reduce their size.
Convert PDF pages to images.
Convert PDFs to Word documents.
Rotate PDF pages.
Add page numbers to PDF files.
Repair corrupted PDF files.
Unlock password-protected PDFs.
Protect PDF files with a password.
Installation
Clone the repository:
sh
Copy code
git clone https://github.com/yourusername/PDFMasterTool.git
Navigate to the project directory:
sh
Copy code
cd PDFMasterTool
Install the required dependencies:
sh
Copy code
pip install -r requirements.txt
Usage
Merging PDFs
sh
Copy code
python pdf_tool.py merge file1.pdf file2.pdf output.pdf
Splitting PDFs
Split by pages per file:
sh
Copy code
python pdf_tool.py split input.pdf output_dir --pages_per_split 5
Split by page range:
sh
Copy code
python pdf_tool.py split input.pdf output_dir --page_range 1-3
Compressing PDFs
sh
Copy code
python pdf_tool.py compress input.pdf output.pdf 60
Converting PDFs to Images
sh
Copy code
python pdf_tool.py pdf_to_images input.pdf output_dir
Converting PDFs to Word
sh
Copy code
python pdf_tool.py pdf_to_word input.pdf output.docx
Rotating PDFs
sh
Copy code
python pdf_tool.py rotate input.pdf output.pdf 90
Adding Page Numbers
sh
Copy code
python pdf_tool.py add_page_numbers input.pdf output.pdf
Repairing PDFs
sh
Copy code
python pdf_tool.py repair input.pdf output.pdf
Unlocking PDFs
sh
Copy code
python pdf_tool.py unlock input.pdf output.pdf password
Protecting PDFs
sh
Copy code
python pdf_tool.py protect input.pdf output.pdf password
Contributing
Contributions are welcome! Please fork the repository and submit a pull request.

License
This project is licensed under the MIT License.
