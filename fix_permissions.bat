@echo off
REM Quick fix script for permission errors on Windows

echo ==========================================
echo Permission Error Fix Script
echo ==========================================
echo.

echo Step 1: Closing Microsoft Word...
taskkill /F /IM WINWORD.EXE >nul 2>&1
if errorlevel 1 (
    echo   No Word processes found (this is good)
) else (
    echo   Word processes closed
)
echo.

echo Step 2: Checking for locked files...
if exist "output-docx\*.docx" (
    echo   Found existing DOCX files
    echo   Attempting to unlock...
    attrib -R "output-docx\*.docx" >nul 2>&1
)
if exist "output-docx\*.pdf" (
    echo   Found existing PDF files
    attrib -R "output-docx\*.pdf" >nul 2>&1
)
echo.

echo Step 3: Optional - Delete existing output files (if you want to start fresh)
echo   This will delete all files in output-docx folder
set /p delete_files="Delete existing output files? (Y/N): "
if /i "%delete_files%"=="Y" (
    echo   Deleting files...
    if exist "output-docx\*.docx" del /F /Q "output-docx\*.docx" >nul 2>&1
    if exist "output-docx\*.pdf" del /F /Q "output-docx\*.pdf" >nul 2>&1
    echo   Files deleted
) else (
    echo   Keeping existing files
)
echo.

echo Step 4: Checking Dropbox status...
echo   If you see Dropbox icon in system tray, make sure it shows "Up to date"
echo   (No files syncing)
echo.

echo ==========================================
echo Next steps:
echo ==========================================
echo.
echo 1. Wait a few seconds for any file locks to clear
echo 2. If using Dropbox, wait until sync is complete
echo 3. Run the script again: python docx_generator.py
echo.
echo If issues persist:
echo - Move project folder outside Dropbox (recommended)
echo - Run Command Prompt as Administrator
echo - See DROPBOX_FIX.md for more solutions
echo.
pause

