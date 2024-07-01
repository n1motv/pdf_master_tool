#!/bin/bash

echo "Updating package lists..."
sudo apt-get update

echo "Installing system dependencies..."
sudo apt-get install -y poppler-utils wkhtmltopdf

echo "Installing Python dependencies..."
pip install ttkthemes PyPDF2 pdf2image pdf2docx pdfkit reportlab

echo "Dependencies installed successfully."
exit

