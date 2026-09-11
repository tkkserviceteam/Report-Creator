CPK Report Generator v0.2 - Build-ready Windows Portable project
================================================================

此資料夾是 Windows x64 Portable EXE 的完整原始專案與自動建置設定。

推薦：打開 BUILD_WITH_GITHUB.txt，透過 GitHub Actions 建置。
這種方式你的 Windows 電腦不需要安裝 Python、PyInstaller、Tesseract。

GitHub 建置完成後，最終使用者只需要：
CPK_Report_Generator.exe

核心功能：
- 客戶名稱 / 機型 / 序號 / 工程師 / 日期
- 預設工程師與機型下拉選單
- CPK1 / CPK2 / Camera Up / Camera Down 圖片
- CPK1 OCR 後預覽與人工修正
- 六項 Cpk 亦可直接手動輸入
- sysdata.txt 最後 100 筆有效資料，取前 6 欄
- 有 Camera Up/Down 使用 sample.xlsx
- 無 Camera Up/Down 使用 sample_without_ccd.xlsx
- 產生 Excel 報告
- 公司章原始物件在插入報告圖片後重新移到 drawing 最上層

目前版本：v0.2 測試版
