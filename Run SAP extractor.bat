@echo off
cd /d "%~dp0"
python sap_web.py
if errorlevel 1 (
  echo.
  echo The app stopped. If packages are missing, run:
  echo     pip install pymupdf pytesseract opencv-python-headless numpy pillow openpyxl flask waitress
  pause
)
