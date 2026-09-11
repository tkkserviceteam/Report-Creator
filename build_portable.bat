@echo off
setlocal
cd /d "%~dp0"
echo ======================================================
echo CPK Report Generator - Local Windows Build
echo ======================================================
echo This method requires Python 3.11 x64 on THIS build PC.
echo For zero-install build, use GitHub Actions (BUILD_WITH_GITHUB.txt).
echo.
python --version >nul 2>&1
if errorlevel 1 (
  echo Python not found.
  pause
  exit /b 1
)
python -m pip install --upgrade pyinstaller pillow
if not exist "ocr\tesseract.exe" (
  echo.
  echo ERROR: ocr\tesseract.exe not found.
  echo Use GitHub Actions to build a package with OCR automatically bundled.
  pause
  exit /b 1
)
call build_github.bat
pause
endlocal
