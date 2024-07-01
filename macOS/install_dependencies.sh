#!/bin/bash

echo "Checking for pip..."

# Check if pip is installed
if ! command -v pip &> /dev/null
then
    echo "pip not found, installing pip..."
    sudo easy_install pip
    sudo pip install --upgrade pip
fi

echo "Updating Homebrew..."
brew update

echo "Installing system dependencies..."
brew install poppler
brew install --cask wkhtmltopdf

echo "Installing Python dependencies..."
pip install ttkthemes PyPDF2 pdf2image pdf2docx pdfkit reportlab

echo "Dependencies installed successfully."
