@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
start "" http://localhost:5000
python app.py
pause
