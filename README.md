# 📦 ExcelVBAWithVbs

這是一套 **Excel / Office VBA 範例模組庫**。  
如果你平常需要整理 Excel 報表、合併檔案、清理資料、匯出 PDF、寄 Outlook 郵件，這個專案可以讓你直接找範例、匯入模組，再依自己的工作流程修改。

目前專案掃描到：

- 📁 **17 個功能分類資料夾**
- 🧩 **437 個 `.bas` VBA 範例模組**
- 🔎 **1 個本地瀏覽器查詢頁面**
- 🧪 **Python + pytest 自動檢查工具**

---

## 🌟 你可以用它做什麼？

- 📊 自動建立 Excel 圖表、樞紐分析表、樞紐分析圖
- 🧹 清理 Excel 資料，例如空白值、重複列、日期格式、百分比、信箱、地址
- 🔍 比對兩份資料、兩張工作表或兩個活頁簿的差異
- 🧮 批次產生 Excel 公式
- 📁 合併或分割 Excel、CSV、文字檔、XML、JSON
- 📄 將工作表、選取範圍、圖表或整本活頁簿匯出 PDF
- 📧 操作 Outlook 郵件、附件、會議邀請、聯絡人與追蹤旗標
- 📝 操作 Word 文件、填入書籤、插入圖表、批次保護與格式設定
- 🖨️ 處理列印、雙面列印、PDF 與檔案工具
- 🗄️ 連接資料庫、讀取組態、產生報表

---

## 🚀 一般使用者最快開始

### 1. 開啟查詢頁面

在專案根目錄打開：

```text
index.html
```

它會自動跳到：

```text
frontend/index.html
```

你可以直接搜尋檔案名稱或 VBA 程式碼內容，例如搜尋：

- `PDF`
- `Outlook`
- `Pivot`
- `Merge`
- `Clean`
- `Attachment`

### 2. 找到需要的 `.bas` 檔

每個 `.bas` 都是一個可匯入 VBA 編輯器的模組。  
建議先看檔名與註解，確認它和你的工作需求接近。

### 3. 匯入 Excel VBA

1. 開啟 Excel。
2. 按 `Alt + F11` 進入 VBA 編輯器。
3. 在左側專案清單中，對目標活頁簿按右鍵。
4. 選擇「匯入檔案」。
5. 選取 `模組/` 底下的 `.bas` 檔。
6. 回到 Excel 執行巨集。

---

## 🗂️ 功能分類速查

| 我想做的事 | 請看這裡 |
|---|---|
| 建立一般 Excel 圖表 | `模組/ChartsNormal` |
| 建立樞紐分析表 | `模組/PivotTableAnalysis` |
| 建立樞紐分析圖 | `模組/PivotCharts` |
| 自動建立公式 | `模組/FormulaCreate` |
| 批次輸入公式 | `模組/BatchEnterFormulas` |
| 條件式格式 | `模組/ConditionalFormatting` |
| 清除儲存格格式 | `模組/ClearCellFormatting` |
| 自動清理資料 | `模組/AutomaticallyCleanData` |
| 自動比對資料差異 | `模組/AutomaticallyCompareDataDifferences` |
| 多條件篩選資料 | `模組/FilterDataBasedonMultipleConditions` |
| 合併多個檔案 | `模組/FileMerge` |
| 分割檔案或工作表 | `模組/FileSplit` |
| 跨工作表合併資料 | `模組/MergeDataAcrossSheets` |
| 匯出 PDF | `模組/ExporttoPDF` |
| Outlook 郵件與行事曆自動化 | `模組/Outlook` |
| Word 文件自動化 | `模組/Word` |
| 通用工具模組 | `模組/Others` |

---

## 📝 Word 範例有哪些？

`模組/Word` 目前包含匯出、合併、格式設定、浮水印、書籤填入、批注擷取等範例，適合想把 Word 文件工作自動化的人。

常見範例包含：

- 📤 將 Excel 資料匯出為 Word 表格
- 🔄 批次取代 Word 文件中的文字
- 🗂️ 合併多個 Word 文件為單一文件
- 📄 批次將 Word 文件轉存為 PDF
- 📋 將 Word 表格匯入 Excel 工作表
- 🔖 以 Excel 資料填入 Word 書籤批次產生文件
- 🏷️ 批次加入頁首頁尾或文字浮水印
- 💬 擷取 Word 批注至 Excel 彙整
- 📊 將 Excel 圖表插入 Word 文件
- 🔐 批次設定或解除 Word 文件密碼保護

---

## 📧 Outlook 範例有哪些？

`模組/Outlook` 目前包含寄信、回覆、附件處理、會議邀請、聯絡人與郵件整理等範例，適合想把 Outlook 日常工作自動化的人。

常見範例包含：

- 📤 建立郵件草稿或寄送郵件
- 📎 儲存所選郵件附件
- 🧾 匯出所選郵件清單為 CSV
- 🗃️ 依主旨關鍵字移動郵件
- 🏷️ 標記郵件後續追蹤
- 💾 將郵件另存為 `.msg`
- 📅 建立行事曆約會或會議邀請草稿
- 👥 從 Excel 資料列建立 Outlook 聯絡人
- 📨 列出目前資料夾未讀郵件

---

## 🧭 專案目錄

```text
ExcelVBAWithVbs/
├── 模組/                 VBA 範例模組主資料夾
├── frontend/             本地查詢頁面與索引資料
├── tests/                自動化檢查程式
├── index.html            查詢頁面入口
├── RunTests.bat          Windows 測試入口
├── RunTests.ps1          PowerShell 測試入口
├── GEN_SAMPLE.md         新增範例時的分類參考
├── AGENTS.md             VBA 編碼與 AI 協作規範
└── README.md             專案說明
```

---

## 🧩 可能需要啟用的 VBA 參考

部分範例會用到 Office 或 Windows 元件。  
如果匯入後出現「使用者定義型態尚未定義」或找不到物件，請到 VBA 編輯器設定：

```text
工具 → 設定引用項目
```

| 使用情境 | 可能需要的參考 |
|---|---|
| Outlook 郵件、行事曆、聯絡人 | Microsoft Outlook xx.0 Object Library |
| Word 文件操作 | Microsoft Word xx.0 Object Library |
| 正規表示式 | Microsoft VBScript Regular Expressions 5.5 |
| 檔案系統工具 | Microsoft Scripting Runtime |
| 資料庫與檔案讀寫 | Microsoft ActiveX Data Objects |
| 網頁擷取 | Microsoft Internet Controls、Microsoft HTML Object Library |
| Acrobat / PDF 進階處理 | Acrobat |

`xx.0` 會依你的 Office 版本不同而變化。

---

## ⚠️ 使用前請先知道

- 💾 執行會修改資料的巨集前，請先備份 Excel 檔。
- 🧪 建議先在測試活頁簿試跑，再套用到正式檔案。
- 🔐 Outlook 自動寄信、刪信、移動郵件等功能，請先確認收件者與資料夾。
- 🌏 VBA 檔案使用 **ANSI / CP950** 編碼，讓繁體中文在 VBA 編輯器正常顯示。
- 🚫 `.bas`、`.cls`、`.frm`、`.frx`、`.vba` 不要改存成 UTF-8 或 UTF-8 BOM。
- 🧱 Sub、Function、變數名稱使用英文；繁體中文只放在註解、字串、工作表名稱或訊息文字中。

---

## 🧪 如何檢查範例是否正常？

### 快速檢查

```batch
RunTests.bat
```

或使用 PowerShell：

```powershell
.\RunTests.ps1
```

### 靜態檢查，不需要 Excel

```powershell
python -m pytest tests/test_static.py -v
```

會檢查：

- ✅ 是否可用 CP950 解碼
- ✅ 是否沒有 UTF-8 BOM
- ✅ 是否有 `Option Explicit`
- ✅ `Sub` / `Function` / `If` / `With` 是否成對
- ✅ 續行符號 ` _` 是否正確
- ✅ Sub / Function 名稱是否避免中文

### Excel 編譯測試，需要桌面版 Excel

```powershell
.\RunTests.ps1 -IncludeExcel
```

第一次執行前，請在 Excel 開啟：

```text
檔案 → 選項 → 信任中心 → 信任中心設定 → 巨集設定
```

並勾選「信任 VBA 專案物件模型的存取」。

---

## 🔄 更新查詢頁面資料

如果新增、刪除或修改 `模組/` 底下的 `.bas` 檔，請重新產生查詢頁面資料：

```powershell
python frontend/build.py
```

如果 Windows 主控台因特殊符號造成輸出錯誤，可改用：

```powershell
$env:PYTHONIOENCODING='utf-8'
python frontend/build.py
```

---

## 👥 適合誰使用？

- 📈 經常整理 Excel 報表的人
- 🧾 需要合併、分割、清理資料的行政、財會或營運人員
- 📧 想減少 Outlook 重複操作的人
- 🧑‍💻 想找 VBA 範例再改成自己工具的開發者
- 🏢 維護企業內部 Office 巨集工具的團隊

---

## 👤 作者

**Dunk (Guan Jhih Liao)**
