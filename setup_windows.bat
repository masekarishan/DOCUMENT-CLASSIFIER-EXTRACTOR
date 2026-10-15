@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
echo.
echo IMPORTANT: Tesseract OCR must also be installed for local image OCR.
echo Default Windows path: C:\Program Files\Tesseract-OCR\tesseract.exe
echo See README.md for details.
pause
