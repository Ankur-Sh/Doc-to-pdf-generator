# Windows Troubleshooting Guide

## PDF Conversion Failed

If you see:
```
Warning: PDF conversion failed for converted_sample_questions_complete.docx. PDF file not created
both docx2pdf and libreoffice conversion methods failed.
```

### This is NOT a PDF viewer problem
This means the conversion from DOCX to PDF failed. The issue is with the conversion tools, not PDF viewers.

### Solutions:

#### 1. Check if LibreOffice is Installed

Run the diagnostic script:
```cmd
python check_libreoffice.py
```

This will tell you if LibreOffice is installed and working.

#### 2. Install LibreOffice (if not installed)

1. Visit: https://www.libreoffice.org/download/
2. Download the Windows version
3. Run the installer
4. **Important:** During installation, make sure to check "Add LibreOffice to PATH" if available
5. Restart Command Prompt completely
6. Run `python check_libreoffice.py` again to verify

#### 3. Add LibreOffice to PATH (if installed but not found)

If LibreOffice is installed but not in PATH:

1. Find LibreOffice installation (usually: `C:\Program Files\LibreOffice\program\`)
2. Copy the full path to `soffice.exe`
3. Add it to Windows PATH:
   - Press `Windows + R`, type `sysdm.cpl`, press Enter
   - Go to "Advanced" tab → "Environment Variables"
   - Under "System Variables", find "Path" → Edit
   - Click "New" → Paste the path (e.g., `C:\Program Files\LibreOffice\program`)
   - Click OK on all windows
   - Restart Command Prompt

#### 4. Test LibreOffice Manually

Try converting a file manually:
```cmd
"C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf "path\to\your\file.docx" --outdir "output\folder"
```

If this works, LibreOffice is installed correctly.

#### 5. Alternative: Use Microsoft Word

If you have Microsoft Word installed:
1. Open the DOCX files in Word
2. File → Save As → Choose PDF format
3. Save the PDFs manually

## Permission Denied Error

If you see an error like:
```
PermissionError: [Errno 13] Permission denied: 'C:\\Users\\...\\output-docx\\converted_sample_questions_only.docx'
```

**Especially if your project is in Dropbox**, this is a common issue.

### Common Causes and Solutions:

#### 1. File is Open in Microsoft Word
**Problem:** The file is currently open in Microsoft Word or another program.

**Solution:**
- Close Microsoft Word completely
- Close any other programs that might be using the file
- Run the script again

#### 2. File is Locked by Windows
**Problem:** Windows has locked the file.

**Solution:**
- Close all programs
- Restart your computer if necessary
- Delete the locked file manually (if it exists)
- Run the script again

#### 3. Insufficient Permissions
**Problem:** You don't have write permissions in the folder.

**Solution:**
- Right-click on the project folder
- Select "Properties"
- Go to "Security" tab
- Make sure your user has "Write" permissions
- Click "Apply" and "OK"

#### 4. Antivirus Software Blocking
**Problem:** Antivirus software is blocking file access.

**Solution:**
- Temporarily disable antivirus
- Add the project folder to antivirus exclusions
- Run the script again

### Quick Fix Steps:

**If your project is in Dropbox (most common cause):**

1. **Run the quick fix script:**
   ```cmd
   fix_permissions.bat
   ```

2. **OR move project outside Dropbox** (BEST SOLUTION):
   - Copy entire folder to: `C:\Users\YourName\Documents\convert text to pdf`
   - Run script from new location
   - See `DROPBOX_FIX.md` for details

3. **OR wait for Dropbox sync:**
   - Check Dropbox icon in system tray
   - Wait until it shows "Up to date"
   - Run script again

**General fixes:**

1. **Close all Word documents**
   - Press `Ctrl + Shift + Esc` to open Task Manager
   - End any "WINWORD.EXE" processes
   - Close all Word windows

2. **Delete existing output files** (if safe to do so)
   ```cmd
   del /F /Q "output-docx\*.docx"
   del /F /Q "output-docx\*.pdf"
   ```

3. **Run as Administrator** (if needed)
   - Right-click on Command Prompt
   - Select "Run as administrator"
   - Navigate to project folder
   - Run the script

4. **Check folder permissions**
   - Right-click project folder → Properties → Security
   - Ensure your user has "Full control" or at least "Write" permission

### Alternative: Use Different Output Folder

If the issue persists, you can modify the script to use a different output folder:

1. Create a new folder (e.g., `C:\MyDocuments\pdf_output`)
2. When prompted for folder path, enter the path to this new folder
3. Make sure you have write permissions in this folder

### Still Having Issues?

1. Check Windows Event Viewer for detailed error messages
2. Make sure you're not running the script from a read-only location (like a CD/DVD)
3. Try running from a different drive (e.g., C:\Users\YourName\Desktop)
4. Check if your antivirus is quarantining the files

