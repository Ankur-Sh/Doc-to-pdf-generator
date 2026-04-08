# Text to PDF Converter

This project converts markdown question files into formatted DOCX and PDF documents for educational content.

## Prerequisites

### 1. Install Python 3.8 or higher

**For macOS:**
```bash
# Option 1: Using Homebrew (recommended)
brew install python3

# Option 2: Download from python.org
# Visit https://www.python.org/downloads/ and download Python 3.x for macOS
```

**For Windows:**
1. Visit https://www.python.org/downloads/
2. Download Python 3.8 or higher
3. Run the installer
4. **Important:** Check "Add Python to PATH" during installation

**For Linux:**
```bash
# Ubuntu/Debian
sudo apt update
sudo apt install python3 python3-pip

# Fedora
sudo dnf install python3 python3-pip
```

### 2. Install LibreOffice (Required for PDF conversion on macOS/Linux)

**For macOS:**
```bash
brew install --cask libreoffice
```

**For Windows:**
1. Visit https://www.libreoffice.org/download/
2. Download and install LibreOffice

**For Linux:**
```bash
# Ubuntu/Debian
sudo apt install libreoffice

# Fedora
sudo dnf install libreoffice
```

**Note:** On Windows, Microsoft Word can be used instead of LibreOffice if you have it installed.

## Installation Steps

### Step 1: Open Terminal/Command Prompt

- **macOS/Linux:** Open Terminal
- **Windows:** Open Command Prompt or PowerShell

### Step 2: Navigate to Project Folder

```bash
cd "/path/to/convert text to pdf"
```

Replace `/path/to/` with the actual path to your project folder.

### Step 3: Install Python Dependencies

```bash
pip3 install -r requirements.txt
```

If `pip3` doesn't work, try:
```bash
python3 -m pip install -r requirements.txt
```

On Windows, you might need:
```bash
pip install -r requirements.txt
```

### Step 4: Verify Installation

```bash
python3 --version
```

You should see Python 3.x.x

## Usage

### Run the Script

```bash
python3 docx_generator.py
```

On Windows, you might need:
```bash
python docx_generator.py
```

### When Prompted:

1. **Generate final PDFs?** 
   - Type `yes` or `y` to generate PDFs
   - Type `no` or `n` to only generate DOCX files

2. **Enter folder path:**
   - Type `.` (period) for current directory, OR
   - Type the full path to your folder containing header.png, footer.png, and first_and_last_page.pdf

3. **Enter font-size:**
   - Type a number (e.g., `12` or `10`)

4. **Enter line-spacing:**
   - Type a number (e.g., `1.5` or `2`)

5. **Enter link (if generating PDFs):**
   - Type the URL to embed in the PDF footer (e.g., `https://example.com/`)

## Project Structure

```
convert text to pdf/
├── files_to_convert/          # Put your .md question files here
│   └── sample_questions.md
├── header.png                  # Header image for documents
├── footer.png                  # Footer image for documents
├── first_and_last_page.pdf     # First and last page template
├── output/                     # Final PDFs (after processing)
├── output-docx/                # DOCX files and intermediate PDFs
├── docx_generator.py           # Main script
├── pdf_transformer.py          # PDF processing script
├── parsing_state.py            # Question parser
├── text_helper.py             # Text formatting helper
└── requirements.txt            # Python dependencies
```

## Supported Input Formats

The script now supports **three input formats**:
- **`.md`** - Markdown files
- **`.docx`** - Microsoft Word documents (newer format)
- **`.doc`** - Microsoft Word documents (older format, requires LibreOffice)

Place your files in the `files_to_convert/` folder.

## Input File Format

Your files (MD, DOCX, or DOC) should follow this format:

```
Question: What is the capital of France?
a) London
b) Berlin
c) Paris
d) Madrid
Answer: c
Explanation: Paris is the capital and largest city of France.
Source: Geography Basics - Chapter 1
```

## Output Files

The script generates 6 different document variants:

1. `questions_only.docx/pdf` - Questions without answers
2. `answer_sheet.docx/pdf` - Answer key table
3. `complete.docx/pdf` - Questions + answers + explanations
4. `explanation_and_answer_only.docx/pdf` - Explanations with answers
5. `explanation_only.docx/pdf` - Explanations only
6. `-classplus.docx` - ClassPlus platform format

## Troubleshooting

### "Python not found"
- Make sure Python is installed and added to PATH
- Try `python3` instead of `python` on macOS/Linux

### "Module not found"
- Run `pip3 install -r requirements.txt` again
- Make sure you're in the project directory

### "PDF conversion failed"
- **macOS/Linux:** Make sure LibreOffice is installed: `brew install --cask libreoffice`
- **Windows:** Make sure LibreOffice or Microsoft Word is installed

### "files_to_convert folder not found"
- Create the folder: `mkdir files_to_convert`
- Add your `.md`, `.doc`, or `.docx` question files to this folder

### "Could not convert DOC to DOCX"
- Install LibreOffice (required for .doc file conversion)
- Or manually convert your .doc files to .docx using Word/LibreOffice

## Quick Start Commands (Copy & Paste)

```bash
# 1. Install Python dependencies
pip3 install -r requirements.txt

# 2. Run the script
python3 docx_generator.py
```

## Support

If you encounter any issues, make sure:
1. Python 3.8+ is installed
2. All dependencies are installed (`pip3 install -r requirements.txt`)
3. LibreOffice is installed (for PDF conversion)
4. The `files_to_convert/` folder exists with `.md` files


