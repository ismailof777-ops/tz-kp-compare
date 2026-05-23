@echo off
cd /d "%~dp0"
title Install OCR for TZ KP Compare
set "BUNDLED_PY=C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

echo Installing OCR Python packages...
if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" -m pip install --target ".python_packages" pymupdf pytesseract
) else (
  py -m pip install --target ".python_packages" pymupdf pytesseract
)

echo.
echo Installing Tesseract OCR...
winget install --id UB-Mannheim.TesseractOCR -e

echo.
echo Downloading Russian OCR language data...
if not exist ".tessdata" mkdir ".tessdata"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri 'https://github.com/tesseract-ocr/tessdata_fast/raw/main/rus.traineddata' -OutFile '.tessdata\rus.traineddata'"

echo.
echo Done. Restart run.bat after installation.
pause
