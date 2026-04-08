@echo off
REM Simple quick fix - just closes Word and waits

echo Closing Microsoft Word...
taskkill /F /IM WINWORD.EXE 2>nul
echo Done! Now wait a few seconds, then run: python docx_generator.py
timeout /t 3 >nul







