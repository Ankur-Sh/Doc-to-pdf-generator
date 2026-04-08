#!/bin/bash
# Script to create a Windows distribution package

echo "=========================================="
echo "Creating Windows Distribution Package"
echo "=========================================="
echo ""

# Get the project folder name
PROJECT_NAME="convert text to pdf"
PARENT_DIR=$(dirname "$(pwd)")
CURRENT_DIR=$(basename "$(pwd)")

# Create package name with timestamp
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
PACKAGE_NAME="convert-text-to-pdf-windows-${TIMESTAMP}"

echo "Project folder: $CURRENT_DIR"
echo "Package name: $PACKAGE_NAME.zip"
echo ""

# Create temporary directory for packaging
TEMP_DIR=$(mktemp -d)
PACKAGE_DIR="$TEMP_DIR/$PACKAGE_NAME"

echo "Creating package structure..."
mkdir -p "$PACKAGE_DIR"

# Copy essential files
echo "Copying files..."

# Python scripts
cp *.py "$PACKAGE_DIR/" 2>/dev/null

# Configuration files
cp requirements.txt "$PACKAGE_DIR/" 2>/dev/null
cp setup.bat "$PACKAGE_DIR/" 2>/dev/null
cp setup.sh "$PACKAGE_DIR/" 2>/dev/null

# Documentation
cp README.md "$PACKAGE_DIR/" 2>/dev/null
cp SETUP_INSTRUCTIONS.txt "$PACKAGE_DIR/" 2>/dev/null
cp WINDOWS_DISTRIBUTION_GUIDE.md "$PACKAGE_DIR/" 2>/dev/null
cp QUICK_START_WINDOWS.txt "$PACKAGE_DIR/" 2>/dev/null
cp WINDOWS_TROUBLESHOOTING.md "$PACKAGE_DIR/" 2>/dev/null
cp CREATE_DISTRIBUTION_PACKAGE.md "$PACKAGE_DIR/" 2>/dev/null
cp *.md "$PACKAGE_DIR/" 2>/dev/null

# Helper scripts
cp *.bat "$PACKAGE_DIR/" 2>/dev/null
cp check_libreoffice.py "$PACKAGE_DIR/" 2>/dev/null

# Optional files (if they exist)
[ -f header.png ] && cp header.png "$PACKAGE_DIR/"
[ -f footer.png ] && cp footer.png "$PACKAGE_DIR/"
[ -f first_and_last_page.pdf ] && cp first_and_last_page.pdf "$PACKAGE_DIR/"

# Create files_to_convert folder
mkdir -p "$PACKAGE_DIR/files_to_convert"
[ -d files_to_convert ] && cp -r files_to_convert/* "$PACKAGE_DIR/files_to_convert/" 2>/dev/null

# Create output folders (empty)
mkdir -p "$PACKAGE_DIR/output"
mkdir -p "$PACKAGE_DIR/output-docx"

# Create README in package
cat > "$PACKAGE_DIR/START_HERE.txt" << 'EOF'
================================================================================
                    TEXT TO PDF CONVERTER - WINDOWS VERSION
================================================================================

QUICK START:
------------
1. Install Python from: https://www.python.org/downloads/
   (Check "Add Python to PATH" during installation)

2. Install LibreOffice from: https://www.libreoffice.org/download/

3. Open Command Prompt in this folder and run:
   setup.bat

4. Then run:
   python docx_generator.py

DETAILED INSTRUCTIONS:
----------------------
- Read: QUICK_START_WINDOWS.txt
- Full guide: WINDOWS_DISTRIBUTION_GUIDE.md
- Troubleshooting: WINDOWS_TROUBLESHOOTING.md

================================================================================
EOF

# Create ZIP file
echo ""
echo "Creating ZIP file..."
cd "$TEMP_DIR"
zip -r "${PACKAGE_NAME}.zip" "$PACKAGE_NAME" -x "*.DS_Store" -x "*__pycache__/*" -x "*.pyc" > /dev/null

# Move ZIP to Desktop
mv "${PACKAGE_NAME}.zip" ~/Desktop/

# Cleanup
rm -rf "$TEMP_DIR"

echo ""
echo "=========================================="
echo "✅ Package created successfully!"
echo "=========================================="
echo ""
echo "Location: ~/Desktop/${PACKAGE_NAME}.zip"
echo ""
echo "You can now share this ZIP file with Windows users."
echo ""






