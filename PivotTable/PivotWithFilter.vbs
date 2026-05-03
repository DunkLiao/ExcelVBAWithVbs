' ============================================================
' PivotWithFilter.vbs
' 說明：使用 VBScript 自動建立含報表篩選頁面的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「業績資料」工作表填入年度銷售示範資料
'   3. 建立樞紐分析表，篩選頁=年度，列=月份，欄=產品，值=銷售額
'   4. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithFilter.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "業績資料"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "年度篩選樞紐"
Const OUTPUT_FILE = "02_PivotWithFilter.xlsx"

Const xlDatabase    = 1
Const xlPageField   = 3
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157

' ── 範例資料（年度、月份、產品、銷售額）──────────────────────
' 2024 年 1-6 月 × 2 產品 = 12 筆，2025 年 1-6 月 × 2 產品 = 12 筆
Dim arrYears(23)
Dim arrMonths(23)
Dim arrProducts(23)
Dim arrAmounts(23)

' 2024 年
arrYears(0)  = 2024 : arrMonths(0)  = 1  : arrProducts(0)  = "筆電" : arrAmounts(0)  = 72000
arrYears(1)  = 2024 : arrMonths(1)  = 1  : arrProducts(1)  = "平板" : arrAmounts(1)  = 43000
arrYears(2)  = 2024 : arrMonths(2)  = 2  : arrProducts(2)  = "筆電" : arrAmounts(2)  = 68000
arrYears(3)  = 2024 : arrMonths(3)  = 2  : arrProducts(3)  = "平板" : arrAmounts(3)  = 39000
arrYears(4)  = 2024 : arrMonths(4)  = 3  : arrProducts(4)  = "筆電" : arrAmounts(4)  = 81000
arrYears(5)  = 2024 : arrMonths(5)  = 3  : arrProducts(5)  = "平板" : arrAmounts(5)  = 51000
arrYears(6)  = 2024 : arrMonths(6)  = 4  : arrProducts(6)  = "筆電" : arrAmounts(6)  = 76000
arrYears(7)  = 2024 : arrMonths(7)  = 4  : arrProducts(7)  = "平板" : arrAmounts(7)  = 47000
arrYears(8)  = 2024 : arrMonths(8)  = 5  : arrProducts(8)  = "筆電" : arrAmounts(8)  = 90000
arrYears(9)  = 2024 : arrMonths(9)  = 5  : arrProducts(9)  = "平板" : arrAmounts(9)  = 58000
arrYears(10) = 2024 : arrMonths(10) = 6  : arrProducts(10) = "筆電" : arrAmounts(10) = 95000
arrYears(11) = 2024 : arrMonths(11) = 6  : arrProducts(11) = "平板" : arrAmounts(11) = 62000
' 2025 年
arrYears(12) = 2025 : arrMonths(12) = 1  : arrProducts(12) = "筆電" : arrAmounts(12) = 88000
arrYears(13) = 2025 : arrMonths(13) = 1  : arrProducts(13) = "平板" : arrAmounts(13) = 54000
arrYears(14) = 2025 : arrMonths(14) = 2  : arrProducts(14) = "筆電" : arrAmounts(14) = 79000
arrYears(15) = 2025 : arrMonths(15) = 2  : arrProducts(15) = "平板" : arrAmounts(15) = 46000
arrYears(16) = 2025 : arrMonths(16) = 3  : arrProducts(16) = "筆電" : arrAmounts(16) = 102000
arrYears(17) = 2025 : arrMonths(17) = 3  : arrProducts(17) = "平板" : arrAmounts(17) = 67000
arrYears(18) = 2025 : arrMonths(18) = 4  : arrProducts(18) = "筆電" : arrAmounts(18) = 96000
arrYears(19) = 2025 : arrMonths(19) = 4  : arrProducts(19) = "平板" : arrAmounts(19) = 59000
arrYears(20) = 2025 : arrMonths(20) = 5  : arrProducts(20) = "筆電" : arrAmounts(20) = 115000
arrYears(21) = 2025 : arrMonths(21) = 5  : arrProducts(21) = "平板" : arrAmounts(21) = 73000
arrYears(22) = 2025 : arrMonths(22) = 6  : arrProducts(22) = "筆電" : arrAmounts(22) = 121000
arrYears(23) = 2025 : arrMonths(23) = 6  : arrProducts(23) = "平板" : arrAmounts(23) = 80000

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objDataSheet, objPivotSheet
Dim objCache, objPivot, objField
Dim savePath, objShell, i

Set objShell = CreateObject("WScript.Shell")
savePath = objShell.SpecialFolders("Desktop") & "\" & OUTPUT_FILE
Set objShell = Nothing

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible       = False
objExcel.DisplayAlerts = False

Set objWorkbook   = objExcel.Workbooks.Add()
Set objDataSheet  = objWorkbook.Sheets(1)
objDataSheet.Name = SHEET_DATA

' ── 寫入標題列 ──────────────────────────────────────────────
objDataSheet.Cells(1, 1).Value = "年度"
objDataSheet.Cells(1, 2).Value = "月份"
objDataSheet.Cells(1, 3).Value = "產品"
objDataSheet.Cells(1, 4).Value = "銷售額"

With objDataSheet.Range("A1:D1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 23
    objDataSheet.Cells(i + 2, 1).Value = arrYears(i)
    objDataSheet.Cells(i + 2, 2).Value = arrMonths(i)
    objDataSheet.Cells(i + 2, 3).Value = arrProducts(i)
    objDataSheet.Cells(i + 2, 4).Value = arrAmounts(i)
Next

objDataSheet.Columns("A:D").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:D25"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定篩選、列、欄、值欄位 ─────────────────────────────────
Set objField = objPivot.PivotFields("年度")
objField.Orientation = xlPageField
objField.Position    = 1

Set objField = objPivot.PivotFields("月份")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("產品")
objField.Orientation = xlColumnField
objField.Position    = 1

Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 銷售額"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "含報表篩選頁面的樞紐分析表：可依年度篩選月份業績"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 14
End With

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objField      = Nothing
Set objPivot      = Nothing
Set objCache      = Nothing
Set objPivotSheet = Nothing
Set objDataSheet  = Nothing
Set objWorkbook   = Nothing
Set objExcel      = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
