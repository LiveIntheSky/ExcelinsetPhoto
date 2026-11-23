# SiC Excel 照片批次自動填入工具 (PhotoToExcelTool)

本工具為專為 SiC 檢驗報告設計的**本機純記憶體執行 · 零落地暫存** Excel 照片批次自動填入桌面網頁工具。

---

## 🚀 快速開始 (雙擊即可運行)

1. **直接雙擊 [run.bat](file:///D:/Desktop/WorkFIle/【Project】SiC插圖補足計畫/run.bat)** 或 **[PhotoToExcelTool.exe](file:///D:/Desktop/WorkFIle/【Project】SiC插圖補足計畫/PhotoToExcelTool.exe)**。
2. 程式啟動時會自動綁定本機動態隨機 Port（例如 `http://127.0.0.1:54718`，防止 Port 衝突），並**自動在系統預設瀏覽器中開啟操作介面**。
3. 在瀏覽器介面中：
   - **步驟 1**：點選「選擇底稿檔案」，選取您的 Excel 檢驗報告底稿（`.xlsx`）。
   - **步驟 2**：點選「選擇照片資料夾」，選取該批次的照片根目錄（例如 `303489-SIC`）。
   - **即時核對**：介面會自動比對白名單與各 Sheet 規則，標示 🟢 命中 與 🔴 該格照片缺失。
   - **步驟 3**：點擊 **「🚀 生成 Excel 報告」**，由後端 ClosedXML 在純記憶體中完成高解析度照片排版與定位，並自動觸發瀏覽器下載完成檔（檔名為 `原檔名_已插圖.xlsx`）。

---

## 📁 檔案結構說明

| 檔案 / 資料夾 | 說明 |
| :--- | :--- |
| 🚀 [**`PhotoToExcelTool.exe`**](file:///D:/Desktop/WorkFIle/【Project】SiC插圖補足計畫/PhotoToExcelTool.exe) | **核心獨立執行檔**（內嵌 .NET 8 執行環境與 ClosedXML 核心，約 49 MB，免安裝直接雙擊運行）。 |
| ⚡ [**`run.bat`**](file:///D:/Desktop/WorkFIle/【Project】SiC插圖補足計畫/run.bat) | 快速啟動捷徑批次檔。 |
| ⚙️ [**`config.yml`**](file:///D:/Desktop/WorkFIle/【Project】SiC插圖補足計畫/config.yml) | 外部設定檔（位於執行檔同級目錄，可隨時用記事本修改 Sheet、尺寸、IQC/FQC 對照）。 |
| 🌐 [**`wwwroot/`**](file:///D:/Desktop/WorkFIle/【Project】SiC插圖補足計畫/wwwroot) | 前端原生 HTML5 / CSS3 / JavaScript 靜態網頁資源。 |
| 📦 [**`PhotoToExcelTool/`**](file:///D:/Desktop/WorkFIle/【Project】SiC插圖補足計畫/PhotoToExcelTool) | 完整 C# 專案原始程式碼。 |
| 📊 `ASM_EPISIL - PE106 CKM Cleaning Report_R302559 - 模板設定.xlsx` | 檢驗報告範本底稿。 |

---

## ⚙️ config.yml 格式架構 (IQC / FQC 獨立模組)

```yaml
accept_extension:
  - ".jpg"
  - ".png"

rules:
  # ----------------------------------------------------------------------------
  # 工作表：Susceptor Sidewalls
  # ----------------------------------------------------------------------------
  - sheet: "Susceptor Sidewalls"
    size:
      width: 13.60        # 照片寬度 (cm)
      height: 10.50       # 照片高度 (cm)
      left_padding: 2.0   # 儲存格左側留白 (mm)
    IQC:
      folder: "PRE/01"    # 清洗前資料夾
      cells:
        1: "B7"
        2: "B28"
        3: "B49"
        4: "B70"
        5: "B91"
        6: "B112"
    FQC:
      folder: "FQC/01"    # 清洗後資料夾
      cells:
        1: "F7"
        2: "F28"
        3: "F49"
        4: "F70"
        5: "F91"
        6: "F112"
```

---

## 🛠️ 開發與重新編譯指令 (若有修改程式碼)

切換至原始碼目錄：
```powershell
cd "D:\Desktop\WorkFIle\【Project】SiC插圖補足計畫\PhotoToExcelTool"
```

- **還原套件**：`dotnet restore`
- **本地偵錯執行**：`dotnet run`
- **重新打包單一執行檔**：
  ```powershell
  dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true -o "D:\Desktop\WorkFIle\【Project】SiC插圖補足計畫\PhotoToExcelTool\publish"
  ```
  發布完成後將 `publish\PhotoToExcelTool.exe` 覆蓋至根目錄即可。
