@echo off
pip install -r requirements.txt
pyinstaller --onefile --windowed --name "PhotoCapture" main.py
echo.
echo Build complete. Find PhotoCapture.exe in the dist\ folder.
pause