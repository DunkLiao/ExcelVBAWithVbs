' ============================================================
' PivotWithDifferenceFrom.vbs
' 說明：使用 VBScript 自動建立顯示差異值的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「年度業績」工作表填入兩年度的銷售示範資料
'   3. 建立樞紐分析表（列=產品，欄=年度，值=銷售額）
'   4. 新增第二個值欄位以「與前一欄差異值」方式顯示年度成長
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithDifferenceFrom.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "年度業績"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "差異值樞紐"
Const OUTPUT_FILE = "13_PivotWithDifferenceFrom.xlsx"

Const xlDatabase        = 1
Const xlRowField        = 1
Const xlColumnField     = 2
Const xlDataField       = 3
Const xlSum             = -4157
Const xlDifferenceFrom  = 2    ' 與基準項差異值

' ── 範例資料（年度、產品線、銷售額）────────────────────────
Dim arrYears(17)
Dim arrLines(17)
Dim arrAmounts(17)

' 2023 年
arrYears(0)  = 2023 : arrLines(0)  = "旗艦機" : arrAmounts(0)  = 285000
arrYears(1)  = 2023 : arrLines(1)  = "標準機" : arrAmounts(1)  = 412000
arrYears(2)  = 2023 : arrLines(2)  = "入門機" : arrAmounts(2)  = 198000
arrYears(3)  = 2023 : arrLines(3)  = "平板"   : arrAmounts(3)  = 156000
arrYears(4)  = 2023 : arrLines(4)  = "穿戴"   : arrAmounts(4)  = 87000
arrYears(5)  = 2023 : arrLines(5)  = "配件"   : arrAmounts(5)  = 63000
' 2024 年
arrYears(6)  = 2024 : arrLines(6)  = "旗艦機" : arrAmounts(6)  = 342000
arrYears(7)  = 2024 : arrLines(7)  = "標準機" : arrAmounts(7)  = 438000
arrYears(8)  = 2024 : arrLines(8)  = "入門機" : arrAmounts(8)  = 175000
arrYears(9)  = 2024 : arrLines(9)  = "平板"   : arrAmounts(9)  = 203000
arrYears(10) = 2024 : arrLines(10) = "穿戴"   : arrAmounts(10) = 128000
arrYears(11) = 2024 : arrLines(11) = "配件"   : arrAmounts(11) = 79000
' 2025 年
arrYears(12) = 2025 : arrLines(12) = "旗艦機" : arrAmounts(12) = 398000
arrYears(13) = 2025 : arrLines(13) = "標準機" : arrAmounts(13) = 462000
arrYears(14) = 2025 : arrLines(14) = "入門機" : arrAmounts(14) = 152000
arrYears(15) = 2025 : arrLines(15) = "平板"   : arrAmounts(15) = 251000
arrYears(16) = 2025 : arrLines(16) = "穿戴"   : arrAmounts(16) = 175000
arrYears(17) = 2025 : arrLines(17) = "配件"   : arrAmounts(17) = 95000

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objDataSheet, objPivotSheet
Dim objCache, objPivot, objField, objDataField
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
objDataSheet.Cells(1, 2).Value = "產品線"
objDataSheet.Cells(1, 3).Value = "銷售額"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 17
    objDataSheet.Cells(i + 2, 1).Value = arrYears(i)
    objDataSheet.Cells(i + 2, 2).Value = arrLines(i)
    objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
Next

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C19"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、欄、值欄位 ──────────────────────────────────────
Set objField = objPivot.PivotFields("產品線")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("年度")
objField.Orientation = xlColumnField
objField.Position    = 1

' 第一個值欄位：一般銷售額
Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "銷售額"

' 第二個值欄位：與前一年度的差異值
Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "年度成長"

' ── 設定差異計算方式（與年度欄位的前一項比較）──────────────
Set objDataField = objPivot.DataFields("年度成長")
objDataField.Calculation = xlDifferenceFrom
objDataField.BaseField   = "年度"
objDataField.BaseItem    = "(previous)"  ' 與前一欄位項目的差異

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "差異值樞紐分析表：同時顯示銷售額及與前一年度的成長差異"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 14
End With

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objDataField  = Nothing
Set objField      = Nothing
Set objPivot      = Nothing
Set objCache      = Nothing
Set objPivotSheet = Nothing
Set objDataSheet  = Nothing
Set objWorkbook   = Nothing
Set objExcel      = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
