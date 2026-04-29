@echo off
cd /d "%~dp0"
echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Building PhotoCapture.exe...
pyinstaller --onefile --windowed --name "PhotoCapture" ^
  --add-data "template\IL 30-1_r6_PT Record.docx;template" ^
  --hidden-import lxml.etree ^
  --hidden-import lxml._elementpath ^
  --hidden-import cv2 ^
  --clean ^
  main.py

echo.
echo Build complete. Find PhotoCapture.exe in the dist\ folder.
pause
