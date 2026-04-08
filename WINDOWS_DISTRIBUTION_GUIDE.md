# Windows Distribution Guide

## How to Share This Project with Windows Users

### Option 1: Share as ZIP File (Recommended)

1. **Create a ZIP file of the project:**
   - Select all project files (except `__pycache__` and `.DS_Store`)
   - Right-click → "Compress" or "Send to → Compressed (zipped) folder"
   - Name it: `convert-text-to-pdf.zip`

2. **Share the ZIP file** via:
   - Email
   - Google Drive / Dropbox
   - USB drive
   - Any file sharing service

3. **Recipient should:**
   - Extract the ZIP file to a folder (e.g., `C:\Users\YourName\Desktop\convert text to pdf`)
   - Follow the setup instructions below

---

## Windows Setup Instructions (For Recipient)

### Step 1: Install Python

1. Visit: **https://www.python.org/downloads/**
2. Download **Python 3.8 or higher** (latest version recommended)
3. Run the installer
4. **IMPORTANT:** ✅ Check **"Add Python to PATH"** checkbox
5. Click **"Install Now"**
6. Wait for installation to complete

**Verify Installation:**
- Open Command Prompt (Press `Win + R`, type `cmd`, press Enter)
- Type: `python --version`
- You should see: `Python 3.x.x`

### Step 2: Install LibreOffice (Required for PDF conversion)

1. Visit: **https://www.libreoffice.org/download/**
2. Download **LibreOffice** for Windows
3. Run the installer
4. Follow the installation wizard (default settings are fine)
5. Complete the installation

**Verify Installation:**
- Open Command Prompt
- Type: `soffice --version`
- You should see version information

### Step 3: Set Up the Project

1. **Extract the project folder** (if you received a ZIP file)
2. **Open Command Prompt** in the project folder:
   - Navigate to the project folder in File Explorer
   - Click in the address bar, type `cmd`, press Enter
   - OR right-click in the folder → "Open in Terminal" / "Open PowerShell window here"

3. **Run the setup script:**
   ```cmd
   setup.bat
   ```
   
   This will automatically:
   - Check if Python is installed
   - Install all required Python packages
   - Verify everything is ready

### Step 4: Prepare Your Files

1. **Create `files_to_convert` folder** (if it doesn't exist)
2. **Add your input files:**
   - `.md` files (markdown format)
   - `.docx` files (Word documents)
   - `.doc` files (older Word format)

3. **Ensure you have these files in the main folder:**
   - `header.png` (optional, for document header)
   - `footer.png` (optional, for document footer)
   - `first_and_last_page.pdf` (optional, for cover pages)

### Step 5: Run the Converter

1. **Open Command Prompt** in the project folder
2. **Run the script:**
   ```cmd
   python docx_generator.py
   ```
3. **Follow the prompts:**
   - Enter "yes" or "no" for PDF generation
   - Enter folder path (or "." for current folder)
   - Enter font size (e.g., 14)
   - Enter line spacing (e.g., 3)
   - Enter footer link (optional)

4. **Find your output files:**
   - DOCX files: `output-docx/` folder
   - PDF files: `output/` folder

---

## Quick Start (If Python is Already Installed)

1. Extract the project folder
2. Open Command Prompt in the project folder
3. Run:
   ```cmd
   setup.bat
   ```
4. Then run:
   ```cmd
   python docx_generator.py
   ```

---

## Troubleshooting

### "Python is not recognized"
- Python is not in PATH
- Reinstall Python and check "Add Python to PATH"
- Or use `py` instead of `python`:
  ```cmd
  py docx_generator.py
  ```

### "Permission denied" errors
- Close Microsoft Word if it's open
- Run `quick_fix.bat` to close Word automatically
- Try running Command Prompt as Administrator

### PDF conversion fails
- Make sure LibreOffice is installed
- Run `check_libreoffice.py` to verify:
  ```cmd
  python check_libreoffice.py
  ```

### File not found errors
- Make sure `files_to_convert` folder exists
- Check that your input files are in the correct folder

---

## Files Included in Distribution

**Required Files:**
- `docx_generator.py` - Main script
- `docx_reader.py` - DOCX/DOC file reader
- `parsing_state.py` - Question parser
- `text_helper.py` - Text formatting utilities
- `pdf_transformer.py` - PDF post-processing
- `table_converter.py` - Table conversion
- `windows_utils.py` - Windows utilities
- `requirements.txt` - Python dependencies
- `setup.bat` - Windows setup script

**Optional Files:**
- `header.png` - Document header image
- `footer.png` - Document footer image
- `first_and_last_page.pdf` - Cover pages
- `README.md` - Documentation
- `WINDOWS_TROUBLESHOOTING.md` - Troubleshooting guide

**Helper Scripts:**
- `quick_fix.bat` - Closes Microsoft Word
- `fix_permissions.bat` - Fixes file permission issues
- `check_libreoffice.py` - Verifies LibreOffice installation

---

## System Requirements

- **Windows 7 or higher**
- **Python 3.8 or higher**
- **LibreOffice** (for PDF conversion)
- **At least 100 MB free disk space**

---

## Support

If you encounter issues:
1. Check `WINDOWS_TROUBLESHOOTING.md`
2. Check `SETUP_INSTRUCTIONS.txt`
3. Run `check_libreoffice.py` to verify LibreOffice
4. Run `quick_fix.bat` if files are locked

---

## Notes

- The project works best when **not** in a Dropbox folder (file locking issues)
- Close Microsoft Word before running the script
- Make sure input files are not open in other programs
- Output files are created in `output-docx/` and `output/` folders






