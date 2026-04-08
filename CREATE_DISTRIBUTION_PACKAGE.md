# How to Create a Distribution Package for Windows Users

## Method 1: Create ZIP File (Easiest)

### On macOS/Linux:

```bash
# Navigate to parent directory
cd ~/Desktop

# Create ZIP excluding unnecessary files
zip -r convert-text-to-pdf.zip "convert text to pdf" \
  -x "*.DS_Store" \
  -x "*__pycache__/*" \
  -x "*.pyc" \
  -x "output/*" \
  -x "output-docx/*" \
  -x "*.log"
```

### On Windows (PowerShell):

```powershell
# Navigate to parent directory
cd $env:USERPROFILE\Desktop

# Create ZIP file
Compress-Archive -Path "convert text to pdf" -DestinationPath "convert-text-to-pdf.zip" -Force
```

### Manual Method:

1. Select the project folder
2. Right-click → "Compress" (macOS) or "Send to → Compressed folder" (Windows)
3. Name it: `convert-text-to-pdf.zip`

## What to Include:

✅ **Include:**
- All `.py` files
- `requirements.txt`
- `setup.bat` and `setup.sh`
- `README.md`
- `SETUP_INSTRUCTIONS.txt`
- `WINDOWS_DISTRIBUTION_GUIDE.md`
- `QUICK_START_WINDOWS.txt`
- `WINDOWS_TROUBLESHOOTING.md`
- `header.png`, `footer.png` (if you have them)
- `first_and_last_page.pdf` (if you have it)
- `files_to_convert/` folder (empty or with sample files)

❌ **Exclude:**
- `__pycache__/` folders
- `.DS_Store` files
- `output/` and `output-docx/` folders (generated files)
- `.git/` folder (if using git)
- Large test files

## Method 2: Share via Cloud Storage

1. Upload the project folder to:
   - Google Drive
   - Dropbox
   - OneDrive
   - GitHub (as a repository)

2. Share the link with the Windows user

3. They download and extract

## Method 3: Create Installer (Advanced)

For a more professional distribution, you could use:
- **PyInstaller** to create a standalone executable
- **Inno Setup** or **NSIS** to create a Windows installer

## Recommended Distribution Package Contents:

```
convert-text-to-pdf/
├── docx_generator.py
├── docx_reader.py
├── parsing_state.py
├── text_helper.py
├── pdf_transformer.py
├── table_converter.py
├── windows_utils.py
├── requirements.txt
├── setup.bat
├── setup.sh
├── quick_fix.bat
├── fix_permissions.bat
├── check_libreoffice.py
├── README.md
├── SETUP_INSTRUCTIONS.txt
├── WINDOWS_DISTRIBUTION_GUIDE.md
├── QUICK_START_WINDOWS.txt
├── WINDOWS_TROUBLESHOOTING.md
├── header.png (optional)
├── footer.png (optional)
├── first_and_last_page.pdf (optional)
└── files_to_convert/
    └── (empty or with sample files)
```

## Instructions for Recipient:

Include this message when sharing:

```
Hi! Here's the Text to PDF Converter project.

QUICK START:
1. Extract the ZIP file
2. Install Python from python.org (check "Add to PATH")
3. Install LibreOffice from libreoffice.org
4. Open Command Prompt in the project folder
5. Run: setup.bat
6. Run: python docx_generator.py

For detailed instructions, see:
- QUICK_START_WINDOWS.txt
- WINDOWS_DISTRIBUTION_GUIDE.md

If you have issues, check WINDOWS_TROUBLESHOOTING.md
```






