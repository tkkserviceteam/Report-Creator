@echo off
setlocal
cd /d "%~dp0"

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist release rmdir /s /q release

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name CPK_Report_Generator ^
  --add-data "config.json;." ^
  --add-data "templates;templates" ^
  --add-data "ocr;ocr" ^
  --hidden-import PIL._tkinter_finder ^
  app.py

if errorlevel 1 exit /b 1

mkdir release
xcopy /E /I /Y "dist\CPK_Report_Generator" "release\CPK_Report_Generator" >nul
if not exist "release\CPK_Report_Generator\output" mkdir "release\CPK_Report_Generator\output"
copy /Y "README_USER.txt" "release\CPK_Report_Generator\README_USER.txt" >nul

echo Build complete.
endlocal
