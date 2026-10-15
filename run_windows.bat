@echo off
setlocal
cd /d "%~dp0"
echo ==============================================
echo Travel Document Extractor

echo Put PDF/JPG/PNG files inside the INPUT folder.
echo Output will be created in OUTPUT\travel_documents.xlsx

echo ==============================================
if not exist input mkdir input
if not exist output mkdir output
python travel_extractor.py --mode auto
if errorlevel 1 (
  echo.
  echo Python returned an error. See output\extraction_log.txt
)
echo.
echo Finished. Check the OUTPUT folder.
pause
