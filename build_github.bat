@echo off
setlocal
cd /d "%~dp0"

if not exist "templates\sample.xlsx" (
  echo ERROR: templates\sample.xlsx is missing from repository root.
  exit /b 2
)
if not exist "templates\sample_without_ccd.xlsx" (
  echo ERROR: templates\sample_without_ccd.xlsx is missing from repository root.
  exit /b 2
)

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist release rmdir /s /q release

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name CPK_Report_Generator ^
  --hidden-import PIL._tkinter_finder ^
  app.py

if errorlevel 1 exit /b 1

mkdir release
xcopy /E /I /Y "dist\CPK_Report_Generator" "release\CPK_Report_Generator" >nul

REM Keep user-editable resources visibly next to the EXE.
xcopy /E /I /Y "templates" "release\CPK_Report_Generator\templates" >nul
copy /Y "config.json" "release\CPK_Report_Generator\config.json" >nul
xcopy /E /I /Y "ocr" "release\CPK_Report_Generator\ocr" >nul
if not exist "release\CPK_Report_Generator\output" mkdir "release\CPK_Report_Generator\output"
copy /Y "README_USER.txt" "release\CPK_Report_Generator\README_USER.txt" >nul

echo Build complete.
endlocal
