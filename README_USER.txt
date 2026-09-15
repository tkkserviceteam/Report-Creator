CPK Report Generator v0.3.1 - Windows Portable

修正：
1. 修正啟動時 ModuleNotFoundError: No module named 'reportlab'。
2. templates 資料夾會直接出現在 CPK_Report_Generator.exe 同一層。
3. config.json、ocr、output 也都放在 EXE 同一層，不需要任何舊版本檔案。

GitHub Actions 建置後，下載 Artifact 並解壓縮。
正常資料夾應包含：
  CPK_Report_Generator.exe
  templates\sample.xlsx
  templates\sample_without_ccd.xlsx
  config.json
  ocr\tesseract.exe
  output\

請整個 CPK_Report_Generator 資料夾一起保留，不要只單獨複製 EXE。
