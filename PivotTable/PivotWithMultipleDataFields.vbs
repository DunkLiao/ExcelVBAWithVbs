' ============================================================
' PivotWithMultipleDataFields.vbs
' 說明：使用 VBScript 自動建立含多個值欄位的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「員工業績」工作表填入員工銷售示範資料
'   3. 建立樞紐分析表，同時顯示銷售額的加總、平均、計數三個值欄位
'   4. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithMultipleDataFields.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "員工業績"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "多值欄位樞紐"
Const OUTPUT_FILE = "04_PivotWithMultipleDataFields.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157
Const xlAverage     = -4106
Const xlCount       = -4112

' ── 範例資料（部門、員工、銷售額）───────────────────────────
Dim arrDepts(19)
Dim arrEmps(19)
Dim arrAmounts(19)

arrDepts(0)  = "業務一部" : arrEmps(0)  = "王小明" : arrAmounts(0)  = 85000
arrDepts(1)  = "業務一部" : arrEmps(1)  = "王小明" : arrAmounts(1)  = 92000
arrDepts(2)  = "業務一部" : arrEmps(2)  = "李大華" : arrAmounts(2)  = 76000
arrDepts(3)  = "業務一部" : arrEmps(3)  = "李大華" : arrAmounts(3)  = 81000
arrDepts(4)  = "業務一部" : arrEmps(4)  = "李大華" : arrAmounts(4)  = 68000
arrDepts(5)  = "業務二部" : arrEmps(5)  = "陳美玲" : arrAmounts(5)  = 102000
arrDepts(6)  = "業務二部" : arrEmps(6)  = "陳美玲" : arrAmounts(6)  = 95000
arrDepts(7)  = "業務二部" : arrEmps(7)  = "陳美玲" : arrAmounts(7)  = 110000
arrDepts(8)  = "業務二部" : arrEmps(8)  = "張志強" : arrAmounts(8)  = 78000
arrDepts(9)  = "業務二部" : arrEmps(9)  = "張志強" : arrAmounts(9)  = 84000
arrDepts(10) = "業務三部" : arrEmps(10) = "林佳慧" : arrAmounts(10) = 63000
arrDepts(11) = "業務三部" : arrEmps(11) = "林佳慧" : arrAmounts(11) = 71000
arrDepts(12) = "業務三部" : arrEmps(12) = "黃文成" : arrAmounts(12) = 88000
arrDepts(13) = "業務三部" : arrEmps(13) = "黃文成" : arrAmounts(13) = 94000
arrDepts(14) = "業務三部" : arrEmps(14) = "黃文成" : arrAmounts(14) = 79000
arrDepts(15) = "業務三部" : arrEmps(15) = "吳雅婷" : arrAmounts(15) = 56000
arrDepts(16) = "業務三部" : arrEmps(16) = "吳雅婷" : arrAmounts(16) = 62000
arrDepts(17) = "業務一部" : arrEmps(17) = "王小明" : arrAmounts(17) = 98000
arrDepts(18) = "業務二部" : arrEmps(18) = "張志強" : arrAmounts(18) = 91000
arrDepts(19) = "業務三部" : arrEmps(19) = "吳雅婷" : arrAmounts(19) = 74000

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
objDataSheet.Cells(1, 1).Value = "部門"
objDataSheet.Cells(1, 2).Value = "員工"
objDataSheet.Cells(1, 3).Value = "銷售額"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 19
    objDataSheet.Cells(i + 2, 1).Value = arrDepts(i)
    objDataSheet.Cells(i + 2, 2).Value = arrEmps(i)
    objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
Next

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C21"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列欄位 ──────────────────────────────────────────────
Set objField = objPivot.PivotFields("部門")
objField.Orientation = xlRowField
objField.Position    = 1

' ── 新增三個值欄位：加總、平均、計數 ────────────────────────
' 加總
Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 銷售額"

' 平均（再次取同欄位並設定不同函數）
Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlAverage
objField.Name        = "平均 - 銷售額"

' 計數
Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlCount
objField.Name        = "計數 - 銷售額"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "多值欄位樞紐分析表：各部門銷售額加總／平均／計數"
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
