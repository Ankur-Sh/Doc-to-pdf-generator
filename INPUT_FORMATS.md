# Supported Input Formats

The converter now supports **multiple input formats** for your question files!

## Supported Formats

### 1. Markdown Files (`.md`)
- **Best for:** Text-based question files
- **Requirements:** None (built-in support)
- **Example:** `sample_questions.md`

### 2. DOCX Files (`.docx`)
- **Best for:** Microsoft Word documents (Word 2007+)
- **Requirements:** `python-docx` library (already included)
- **Example:** `questions.docx`

### 3. DOC Files (`.doc`)
- **Best for:** Older Microsoft Word documents (Word 97-2003)
- **Requirements:** LibreOffice installed (for conversion)
- **Example:** `old_questions.doc`

## How to Use

1. **Place your files** in the `files_to_convert/` folder
2. **Run the script:** `python docx_generator.py`
3. **The script will automatically:**
   - Detect the file format
   - Extract text from DOC/DOCX files
   - Parse questions using the same format
   - Generate formatted documents

## File Format Requirements

All formats must follow the same structure:

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

## DOC File Conversion

If you have `.doc` files (older format):
- The script will **automatically convert** them using LibreOffice
- Make sure LibreOffice is installed
- If conversion fails, manually convert to DOCX first

## Tips

1. **For best results:** Use `.md` or `.docx` files
2. **For old documents:** Convert `.doc` to `.docx` first if possible
3. **File structure:** Keep the same question format regardless of file type
4. **Multiple files:** You can mix `.md`, `.docx`, and `.doc` files in the same folder

## Examples

```
files_to_convert/
├── questions.md          ✓ Supported
├── questions.docx        ✓ Supported
├── old_questions.doc     ✓ Supported (needs LibreOffice)
└── questions.txt         ✗ Not supported
```

All supported files will be processed automatically!






