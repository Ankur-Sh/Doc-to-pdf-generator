#!/bin/bash
# Setup script for Text to PDF Converter

echo "=========================================="
echo "Text to PDF Converter - Setup Script"
echo "=========================================="
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed!"
    echo ""
    echo "Please install Python 3 first:"
    echo "  macOS: brew install python3"
    echo "  Or visit: https://www.python.org/downloads/"
    exit 1
fi

echo "✓ Python 3 found: $(python3 --version)"
echo ""

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip3 is not installed!"
    echo "Please install pip3 first."
    exit 1
fi

echo "✓ pip3 found"
echo ""

# Install dependencies
echo "Installing Python dependencies..."
pip3 install -r requirements.txt

if [ $? -eq 0 ]; then
    echo ""
    echo "✓ Dependencies installed successfully!"
    echo ""
    echo "Next steps:"
    echo "1. Make sure you have header.png, footer.png, and first_and_last_page.pdf in this folder"
    echo "2. Create a 'files_to_convert' folder and add your .md question files"
    echo "3. Run: python3 docx_generator.py"
    echo ""
else
    echo ""
    echo "❌ Failed to install dependencies"
    echo "Please check the error messages above"
    exit 1
fi







