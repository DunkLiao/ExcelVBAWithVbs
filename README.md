# ExcelVBAWithVbs

這是一個用 **VBS + Excel VBA** 做報表與資料處理自動化的範例專案。

目前已實作一條完整參考流程：

1. 用 VBS 建立一個啟用巨集的 Excel 活頁簿
2. 匯入 repo 內的 VBA 模組
3. 載入 CSV 資料
4. 由 VBA 產生彙總報表
5. 將活頁簿中的 VBA 模組匯回 repo

## 專案結構

```text
src\vba\                 VBA 原始碼（版控主體）
scripts\                 VBS 自動化腳本
sample-data\             範例輸入資料
workbooks\               執行時產生的 Excel 活頁簿
```

目前的主要檔案：

- `src\vba\modules\ReportAutomation.bas`：報表產生邏輯
- `scripts\bootstrap-report-workbook.vbs`：建立 `.xlsm` 活頁簿並匯入 VBA 模組
- `scripts\run-sales-report.vbs`：開啟活頁簿、載入 CSV、執行 VBA 報表流程
- `scripts\export-vba-modules.vbs`：把活頁簿中的 VBA 模組匯回 repo
- `sample-data\sales-input.csv`：範例資料

## 運作方式

這個 repo 採用 **原始碼 / 產物分離** 的方式：

- `src\vba\` 裡的匯出模組是主要原始碼
- `workbooks\*.xlsm` 是執行時活頁簿產物
- `scripts\*.vbs` 負責用 Excel COM 把兩邊接起來

也就是說，平常應該把邏輯維護在匯出的 `.bas/.cls/.frm` 檔案，而不是只依賴活頁簿內的 VBA。

## 環境需求

- Windows
- 已安裝 Excel Desktop
- Excel 啟用 **Trust access to the VBA project object model**

如果沒有開啟 VBA 專案存取權限，建立活頁簿或匯出模組的腳本會失敗。

## 快速開始

### 1. 建立活頁簿

```powershell
cscript //nologo scripts\bootstrap-report-workbook.vbs
```

預設會建立：

```text
workbooks\ReportAutomationTemplate.xlsm
```

### 2. 執行範例報表

```powershell
cscript //nologo scripts\run-sales-report.vbs sample-data\sales-input.csv
```

這會：

- 將 CSV 載入 `InputData`
- 依 `Team` 與 `Amount` 欄位彙總資料
- 將結果寫到 `Report`

### 3. 將 VBA 匯回 repo

如果你在 Excel 裡修改了 VBA，再執行：

```powershell
cscript //nologo scripts\export-vba-modules.vbs
```

## 目前的報表流程

`ReportAutomation.LoadCsvAndGenerateReport` 目前假設輸入 CSV 至少包含以下欄位：

```text
Team,Amount
```

範例輸入檔在：

```text
sample-data\sales-input.csv
```

輸出結果會寫入活頁簿中的：

- `InputData`
- `Report`

## 驗證方式

這個專案目前沒有獨立的 build、lint 或自動化測試框架。實際驗證方式就是跑完整腳本流程：

```powershell
cscript //nologo scripts\bootstrap-report-workbook.vbs
cscript //nologo scripts\run-sales-report.vbs sample-data\sales-input.csv
cscript //nologo scripts\export-vba-modules.vbs
```

如果只想做最小驗證，可直接跑：

```powershell
cscript //nologo scripts\run-sales-report.vbs sample-data\sales-input.csv
```

## 開發慣例

- 新的 VBA 邏輯優先放在 `src\vba\` 的匯出模組
- 新的 VBS 腳本沿用目前做法，從 `WScript.ScriptFullName` 推出 `repoRoot`
- 預設活頁簿名稱沿用 `workbooks\ReportAutomationTemplate.xlsm`
- 發生檔案、Excel、欄位結構錯誤時，維持目前 `Err.Raise` 的明確錯誤處理方式

## 目前限制

- `scripts\export-vba-modules.vbs` 目前只匯出 standard/class/form modules
- `ThisWorkbook` 或工作表事件模組尚未納入匯出流程
- 如果之後要加入事件驅動 VBA，需要一起調整匯出策略
