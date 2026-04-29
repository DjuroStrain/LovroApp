@echo off
cd /d "%~dp0"
echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Building PhotoCapture.exe...
pyinstaller PhotoCapture.spec --clean

echo.
echo Build complete. Find PhotoCapture.exe in the dist\ folder.
pause