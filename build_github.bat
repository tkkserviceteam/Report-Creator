@echo off
setlocal
cd /d "%~dp0"
if not exist "templates\sample.xlsx" exit /b 2
if not exist "templates\sample_without_ccd.xlsx" exit /b 2
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist release rmdir /s /q release
python -m PyInstaller --noconfirm --clean --onefile --windowed --name CPK_Report_Generator --hidden-import PIL._tkinter_finder app.py
if errorlevel 1 exit /b 1
mkdir "release\CPK_Report_Generator"
copy /Y "dist\CPK_Report_Generator.exe" "release\CPK_Report_Generator\CPK_Report_Generator.exe" >nul
xcopy /E /I /Y "templates" "release\CPK_Report_Generator\templates" >nul
copy /Y "config.json" "release\CPK_Report_Generator\config.json" >nul
xcopy /E /I /Y "ocr" "release\CPK_Report_Generator\ocr" >nul
if not exist "release\CPK_Report_Generator\Report History" mkdir "release\CPK_Report_Generator\Report History"
copy /Y "README_USER.txt" "release\CPK_Report_Generator\README_USER.txt" >nul
echo Build complete.
endlocal
