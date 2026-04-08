@echo off
REM Setup script for Text to PDF Converter (Windows)

echo ==========================================
echo Text to PDF Converter - Setup Script
echo ==========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed!
    echo.
    echo Please install Python 3 first:
    echo Visit: https://www.python.org/downloads/
    echo.
    echo IMPORTANT: Check "Add Python to PATH" during installation!
    pause
    exit /b 1
)

echo [OK] Python found
python --version
echo.

REM Check if pip is installed
pip --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pip is not installed!
    echo Please install pip first.
    pause
    exit /b 1
)

echo [OK] pip found
echo.

REM Install dependencies
echo Installing Python dependencies...
pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install dependencies
    echo Please check the error messages above
    pause
    exit /b 1
)

echo.
echo [OK] Dependencies installed successfully!
echo.
echo Next steps:
echo 1. Make sure you have header.png, footer.png, and first_and_last_page.pdf in this folder
echo 2. Create a 'files_to_convert' folder and add your .md question files
echo 3. Run: python docx_generator.py
echo.
pause







