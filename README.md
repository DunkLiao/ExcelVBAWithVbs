# ExcelVBAWithVbs

Excel / Office VBA 模組工具庫，收錄可直接匯入 VBA 編輯器使用的 `.bas` / `.cls` / `.frm` 模組。

## 專案簡介

本儲存庫專門收錄 Microsoft Office VBA 程式碼，涵蓋 Excel、Outlook 自動化範例與通用工具模組。所有模組均以 **ANSI / CP950** 編碼儲存，可直接匯入 VBA 編輯器並通過編譯。

適合用於：

- Excel 報表自動化
- Outlook 郵件處理
- 資料庫存取（ODBC / SQL Server / DB2 / DBF）
- 批次資料整理與格式轉換
- 企業內部流程工具開發

---

## 目錄結構

```
模組/
├── ChartsNormal/                         — 圖表建立範例（長條、折線、圓餅、圓環等）
├── PivotTableAnalysis/                   — 樞紐分析表操作範例
├── PivotCharts/                          — 樞紐分析圖範例
├── FormulaCreate/                        — 各類公式建立範例（陣列、財務、統計等）
├── ConditionalFormatting/                — 條件式格式設定範例
├── ClearCellFormatting/                  — 清除儲存格格式範例
├── BatchEnterFormulas/                   — 批次輸入公式範例
├── FileMerge/                            — 合併多個 Excel 檔案範例
├── FileSplit/                            — 依條件分割工作表範例
├── MergeDataAcrossSheets/                — 跨工作表彙整資料範例
├── ExporttoPDF/                          — 匯出工作表 / 活頁簿為 PDF 範例
├── FilterDataBasedonMultipleConditions/  — 多條件篩選範例
├── AutomaticallyCleanData/              — 自動清理資料範例
├── AutomaticallyCompareDataDifferences/ — 自動比對差異範例
├── Outlook/                              — Outlook 自動化模組
├── 組態讀取/                             — 報表組態讀取模組
├── 個人工作活頁簿/                       — 個人活頁簿自訂模組
├── 視窗/                                 — UserForm 表單範本
└── *.bas                                 — 通用工具模組（資料庫、檔案、字串、列印等）
```

---

## 如何使用

### 匯入模組

1. 開啟 Excel，按 **Alt + F11** 進入 VBA 編輯器。
2. 在「專案」面板中，對目標活頁簿按右鍵 → **匯入檔案**。
3. 選取 `.bas` 或 `.cls` 檔案即可匯入。

### 常用模組快速參考

| 需求 | 位置 |
|------|------|
| 傳送 Outlook 郵件 / 電子報 | `Outlook/` 資料夾 |
| 樞紐分析表建立與操作 | `PivotTableAnalysis/` 資料夾 |
| 樞紐分析圖 | `PivotCharts/` 資料夾 |
| 圖表建立範例 | `ChartsNormal/` 資料夾 |
| 各類公式建立 | `FormulaCreate/` 資料夾 |
| 條件式格式設定 | `ConditionalFormatting/` 資料夾 |
| 清除儲存格格式 | `ClearCellFormatting/` 資料夾 |
| 批次輸入公式 | `BatchEnterFormulas/` 資料夾 |
| 合併多個 Excel 檔案 | `FileMerge/` 資料夾 |
| 分割工作表 | `FileSplit/` 資料夾 |
| 跨工作表彙整資料 | `MergeDataAcrossSheets/` 資料夾 |
| 匯出 PDF | `ExporttoPDF/` 資料夾 |
| 多條件篩選 | `FilterDataBasedonMultipleConditions/` 資料夾 |
| 自動清理資料 | `AutomaticallyCleanData/` 資料夾 |
| 自動比對差異 | `AutomaticallyCompareDataDifferences/` 資料夾 |
| 報表組態讀取 | `組態讀取/` 資料夾 |
| UserForm 表單範本 | `視窗/` 資料夾 |
| 資料庫 / 檔案 / 字串 / 列印工具 | `模組/*.bas` 通用工具模組 |

---

## 外部參考 (References)

部分模組需在 VBA 編輯器中啟用對應的外部參考（**工具 → 設定引用項目**）：

| 模組 / 資料夾 | 需要的 Reference |
|------|-----------------|
| `Outlook/` 資料夾 | Microsoft Outlook xx.0 Object Library |
| 列印 / PDF 相關模組 | Acrobat |
| 正規表達式相關模組 | Microsoft VBScript Regular Expressions 5.5 |
| ODBC 相關模組 | 使用 Windows API，無需額外 Reference |
| 網頁擷取相關模組 | Microsoft Internet Controls、Microsoft HTML Object Library |
| 資料庫 / 檔案讀寫相關模組 | Microsoft ActiveX Data Objects (ADODB) |

---

## 自動化測試

本專案內建兩階段自動化測試，使用 Python + pytest 執行。

### 快速執行

```batch
RunTests.bat
```

或使用 PowerShell：

```powershell
.\RunTests.ps1               # Phase 1 靜態檢查
.\RunTests.ps1 -IncludeExcel # Phase 1 + Phase 2（需桌面版 Excel）
.\RunTests.ps1 -Verbose      # 顯示詳細輸出
```

### Phase 1：靜態檢查（不需要 Excel）

對所有 `.bas` 檔案進行靜態分析，不需要安裝 Excel：

| 檢查項目 | 說明 |
|----------|------|
| CP950 編碼 | 檔案必須能以 CP950 解碼 |
| 無 UTF-8 BOM | 前三個 bytes 不得為 `\xef\xbb\xbf` |
| Option Explicit | 每個模組第一個程式碼行必須是 `Option Explicit` |
| Sub / End Sub 成對 | `Sub` 與 `End Sub` 數量相符 |
| Function / End Function 成對 | `Function` 與 `End Function` 數量相符 |
| With / End With 成對 | `With` 與 `End With` 數量相符 |
| If / End If 成對 | 多行 `If...Then` 必須有對應 `End If` |
| 續行符號空格 | `_` 前方必須有空格 |
| Sub / Function 名稱 | 不得含繁體或簡體中文字元 |

```bash
python -m pytest tests/test_static.py -v
```

### Phase 2：Excel COM 動態編譯（需桌面版 Excel）

透過 Excel COM 物件將 `.bas` 匯入暫存活頁簿，觸發 VBE `Compile VBAProject` 指令：

**前置條件：**
> Excel → 檔案 → 選項 → 信任中心 → 信任中心設定 → 巨集設定  
> 勾選「**信任 VBA 專案物件模型的存取**」

```bash
pip install pywin32
python -m pytest tests/test_compile.py -v
```

---

## 編碼規範

- 所有 `.bas` / `.cls` / `.frm` 檔案以 **ANSI / CP950** 編碼儲存。
- 不使用 UTF-8 或 UTF-8 BOM，以確保 VBA 編輯器正確顯示繁體中文。
- `Sub` / `Function` 名稱使用英文；繁體中文僅出現於註解、字串、工作表名稱。
- 每個模組第一行使用 `Option Explicit`，所有變數需明確宣告。

---

## 開發規範

- 每個 `If` / `For` / `With` / `Sub` / `Function` 必須有對應的結尾關鍵字。
- 續行符號 ` _` 前方必須有空格。
- 不使用全形標點符號於 VBA 語法中。
- 繁體中文僅限出現於：註解、字串、`MsgBox`、工作表名稱、圖表標題。
- 具備錯誤處理（`On Error GoTo` / `On Error Resume Next`）。

---

## 作者

**Dunk (Guan Jhih Liao)**