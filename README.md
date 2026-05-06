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
├── 資料庫 / SQL
│   ├── ADQuery.bas                — 查詢 Active Directory 使用者資訊
│   ├── ConnectDBF.bas             — 連接 DBF 資料庫
│   ├── ConnectionToOutput.bas     — 資料庫查詢結果輸出至工作表
│   ├── ODBCUtility.bas            — 列舉系統 ODBC 資料來源 (DSN)
│   ├── SqlStatement.bas           — SQL Server 查詢語句產生器
│   └── SqlStatementDB2.bas        — DB2 查詢語句產生器
│
├── 工作表 / 活頁簿操作
│   ├── CellFunction.bas           — 儲存格常用函式 (取得最後列/欄等)
│   ├── CheckFormat.bas            — 格式驗證工具
│   ├── SaveWorkSheetTool.bas      — 工作表另存工具
│   ├── SetSheetPureText.bas       — 將工作表資料轉換為純文字
│   ├── SheetToolUtil.bas          — 工作表工具集
│   ├── SheetUtil.bas              — 工作表通用操作 (自動調整欄寬/列高等)
│   ├── SortUtility.bas            — 資料排序工具
│   └── 活頁簿管理.bas             — 活頁簿與工作表管理
│
├── 檔案 / 路徑操作
│   ├── CombineExeFileUtil.bas     — 合併執行檔工具
│   ├── CopyFilePath.bas           — 複製檔案路徑工具
│   ├── FileInfo.bas               — 取得檔案資訊
│   ├── FileIOUtility.bas          — 檔案讀寫工具 (含 ADODB.Stream 編碼支援)
│   └── ShellExecTool.bas          — 呼叫 Shell 命令工具
│
├── 列印 / PDF
│   ├── PdfUtility.bas             — PDF 合併 / 分割 (需 Acrobat 參考)
│   ├── PrintUtil.bas              — 列印工具 (批次列印、列印至 PDF)
│   ├── SetPrinterDuplex.bas       — 設定印表機雙面列印
│   └── SettingPrinterDuplexNew.bas— 雙面列印設定 (新版)
│
├── 郵件 / Outlook
│   ├── SendingMail.bas            — 透過 Outlook 傳送 HTML 郵件
│   ├── SendNews.bas               — 批次發送電子報
│   └── Outlook/
│       ├── AutoReply.bas          — Outlook 自動回覆規則
│       ├── DealOilFunAttachment.bas — 處理特定附件
│       ├── PrintAttachment.bas    — 列印附件
│       ├── RemoveMail.bas         — 批次刪除郵件
│       ├── SaveFundFiles.bas      — 儲存附件至指定資料夾
│       ├── SendMail.bas           — Outlook 傳送郵件
│       ├── ThisOutlookSession.cls — Outlook Session 事件模組
│       └── UnzipFile.bas          — 解壓縮附件
│
├── 字串 / 編碼工具
│   ├── EncodingUtil.bas           — 偵測與轉換檔案編碼
│   ├── RegExpTool.bas             — 正規表達式工具 (需 VBScript RegExp 5.5)
│   └── StringUtility.bas          — 字串處理工具 (日期轉換、位元組截字等)
│
├── 其他工具
│   ├── CopyCliboardUtility.bas    — 剪貼簿工具
│   ├── CustmizeForm.bas           — 自訂表單工具
│   ├── LoadConfigFile.bas         — 讀取組態設定檔 (INI 格式)
│   ├── LookMutiInCell.bas         — 多值儲存格查詢
│   ├── PasswordRemoveUtil.bas     — 移除工作表保護密碼
│   ├── SavingPictureTool.bas      — 圖片儲存工具
│   ├── ShapUtil.bas               — 圖案 (Shape) 操作工具
│   └── WebHtmlFetch.bas           — 擷取網頁 URL 與 HTML 內容
│
├── charts/                        — 圖表建立範例 (10 種圖表類型)
│   ├── AreaChartExample.bas       — 區域圖
│   ├── BarChartExample.bas        — 長條圖
│   ├── BubbleChartExample.bas     — 泡泡圖
│   ├── ComboChartExample.bas      — 組合圖
│   ├── LineChartExample.bas       — 折線圖
│   ├── PieChartExample.bas        — 圓餅圖
│   ├── RadarChartExample.bas      — 雷達圖
│   ├── ScatterChartExample.bas    — 散佈圖
│   ├── StockChartExample.bas      — 股票圖
│   └── SurfaceChartExample.bas    — 曲面圖
│
├── PivotTable/                    — 樞紐分析表範例 (20 種操作情境)
│   ├── BasicPivotTable.bas        — 基本樞紐分析表建立
│   ├── CalculatedFieldPivot.bas   — 計算欄位
│   ├── ConditionalFormatPivot.bas — 條件式格式設定
│   ├── CopyPivotAsValues.bas      — 樞紐複製為純值
│   ├── DateGroupPivotTable.bas    — 日期群組
│   ├── DrillDownPivotTable.bas    — 向下切入 (Drill Down)
│   ├── FilterPivotTable.bas       — 報表篩選器
│   ├── HideItemPivot.bas          — 隱藏項目
│   ├── LayoutPivotTable.bas       — 版面配置設定
│   ├── MultiFieldPivotTable.bas   — 多欄位樞紐
│   ├── PercentagePivotTable.bas   — 百分比顯示
│   ├── PivotChartCreate.bas       — 樞紐圖建立
│   ├── RankPivotTable.bas         — 排名顯示
│   ├── RefreshPivotTable.bas      — 重新整理
│   ├── RunningTotalPivot.bas      — 累計加總
│   ├── SlicerPivotTable.bas       — 交叉分析篩選器
│   ├── SortPivotTable.bas         — 排序
│   ├── SubtotalControlPivot.bas   — 小計控制
│   ├── TimelinePivot.bas          — 時間表篩選器
│   └── TopNFilterPivot.bas        — 前 N 名篩選
│
├── 組態讀取/                      — 報表組態讀取範例
│   ├── LoadConfigFile.bas         — 組態檔讀取
│   ├── ODBCUtility.bas            — ODBC 工具（組態版）
│   ├── FileSave.bas               — 檔案儲存邏輯
│   ├── GenReportA.bas             — 報表 A 產生器
│   ├── GenReportM.bas             — 報表 M 產生器
│   └── GenReportQ.bas             — 報表 Q 產生器
│
├── 個人工作活頁簿/                — 個人活頁簿自訂模組
│   ├── CustmizeForm.bas           — 自訂表單
│   └── SheetUtil.bas              — 工作表工具
│
└── 視窗/                          — UserForm 表單範本
    ├── frmTemplate.frm / .frx     — 表單範本 1
    └── frmTemplate2.frm / .frx    — 表單範本 2
```

---

## 如何使用

### 匯入模組

1. 開啟 Excel，按 **Alt + F11** 進入 VBA 編輯器。
2. 在「專案」面板中，對目標活頁簿按右鍵 → **匯入檔案**。
3. 選取 `.bas` 或 `.cls` 檔案即可匯入。

### 常用模組快速參考

| 需求 | 模組 |
|------|------|
| 傳送 Outlook 郵件 | `SendingMail.bas` |
| 批次發電子報 | `SendNews.bas` |
| 檔案讀寫（含編碼控制） | `FileIOUtility.bas` |
| 正規表達式 | `RegExpTool.bas` |
| 字串處理 / 日期轉換 | `StringUtility.bas` |
| SQL Server 查詢語句產生 | `SqlStatement.bas` |
| DB2 查詢語句產生 | `SqlStatementDB2.bas` |
| ODBC 資料來源列舉 | `ODBCUtility.bas` |
| AD 使用者查詢 | `ADQuery.bas` |
| DBF 資料庫連接 | `ConnectDBF.bas` |
| PDF 合併 / 分割 | `PdfUtility.bas` |
| 批次列印 / 列印至 PDF | `PrintUtil.bas` |
| 雙面列印設定 | `SettingPrinterDuplexNew.bas` |
| 工作表操作 | `SheetUtil.bas` |
| 儲存格常用函式 | `CellFunction.bas` |
| 樞紐分析表（20 種情境） | `PivotTable/` 資料夾 |
| 圖表建立範例（10 種） | `charts/` 資料夾 |
| Outlook 自動化 | `Outlook/` 資料夾 |
| 網頁內容擷取 | `WebHtmlFetch.bas` |
| UserForm 表單範本 | `視窗/` 資料夾 |

---

## 外部參考 (References)

部分模組需在 VBA 編輯器中啟用對應的外部參考（**工具 → 設定引用項目**）：

| 模組 | 需要的 Reference |
|------|-----------------|
| `SendingMail.bas` | Microsoft Outlook xx.0 Object Library |
| `Outlook/` 資料夾 | Microsoft Outlook xx.0 Object Library |
| `PdfUtility.bas` | Acrobat |
| `RegExpTool.bas` | Microsoft VBScript Regular Expressions 5.5 |
| `ODBCUtility.bas` | 使用 Windows API，無需額外 Reference |
| `WebHtmlFetch.bas` | Microsoft Internet Controls、Microsoft HTML Object Library |
| `FileIOUtility.bas` | Microsoft ActiveX Data Objects (ADODB) |
| `ConnectDBF.bas` | Microsoft ActiveX Data Objects (ADODB) |

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