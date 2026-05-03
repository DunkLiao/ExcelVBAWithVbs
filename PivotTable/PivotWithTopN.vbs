' ============================================================
' PivotWithTopN.vbs
' 說明：使用 VBScript 自動建立含前 N 名篩選的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「業務業績」工作表填入業務員銷售示範資料
'   3. 建立樞紐分析表，並篩選出銷售額前 5 名的業務員
'   4. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithTopN.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "業務業績"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "前5名樞紐"
Const OUTPUT_FILE = "07_PivotWithTopN.xlsx"
Const TOP_N       = 5   ' 顯示前幾名

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlDataField   = 3
Const xlSum         = -4157
Const xlAutomatic   = -4105  ' AutoShow Type：自動
Const xlTop         = 1      ' AutoShow Range：前幾名

' ── 範例資料（業務員、季度、銷售額）────────────────────────
Dim arrEmps(27)
Dim arrQtrs(27)
Dim arrAmounts(27)

' 8 位業務員 × 4 季 = 32 筆（此處示範 28 筆，缺部分季度資料）
arrEmps(0)  = "王小明" : arrQtrs(0)  = "Q1" : arrAmounts(0)  = 85000
arrEmps(1)  = "王小明" : arrQtrs(1)  = "Q2" : arrAmounts(1)  = 92000
arrEmps(2)  = "王小明" : arrQtrs(2)  = "Q3" : arrAmounts(2)  = 78000
arrEmps(3)  = "王小明" : arrQtrs(3)  = "Q4" : arrAmounts(3)  = 105000
arrEmps(4)  = "李大華" : arrQtrs(4)  = "Q1" : arrAmounts(4)  = 63000
arrEmps(5)  = "李大華" : arrQtrs(5)  = "Q2" : arrAmounts(5)  = 71000
arrEmps(6)  = "李大華" : arrQtrs(6)  = "Q3" : arrAmounts(6)  = 58000
arrEmps(7)  = "李大華" : arrQtrs(7)  = "Q4" : arrAmounts(7)  = 80000
arrEmps(8)  = "陳美玲" : arrQtrs(8)  = "Q1" : arrAmounts(8)  = 112000
arrEmps(9)  = "陳美玲" : arrQtrs(9)  = "Q2" : arrAmounts(9)  = 128000
arrEmps(10) = "陳美玲" : arrQtrs(10) = "Q3" : arrAmounts(10) = 135000
arrEmps(11) = "陳美玲" : arrQtrs(11) = "Q4" : arrAmounts(11) = 149000
arrEmps(12) = "張志強" : arrQtrs(12) = "Q1" : arrAmounts(12) = 77000
arrEmps(13) = "張志強" : arrQtrs(13) = "Q2" : arrAmounts(13) = 84000
arrEmps(14) = "張志強" : arrQtrs(14) = "Q3" : arrAmounts(14) = 91000
arrEmps(15) = "張志強" : arrQtrs(15) = "Q4" : arrAmounts(15) = 99000
arrEmps(16) = "林佳慧" : arrQtrs(16) = "Q1" : arrAmounts(16) = 54000
arrEmps(17) = "林佳慧" : arrQtrs(17) = "Q2" : arrAmounts(17) = 61000
arrEmps(18) = "林佳慧" : arrQtrs(18) = "Q3" : arrAmounts(18) = 49000
arrEmps(19) = "林佳慧" : arrQtrs(19) = "Q4" : arrAmounts(19) = 72000
arrEmps(20) = "黃文成" : arrQtrs(20) = "Q1" : arrAmounts(20) = 98000
arrEmps(21) = "黃文成" : arrQtrs(21) = "Q2" : arrAmounts(21) = 107000
arrEmps(22) = "黃文成" : arrQtrs(22) = "Q3" : arrAmounts(22) = 118000
arrEmps(23) = "黃文成" : arrQtrs(23) = "Q4" : arrAmounts(23) = 132000
arrEmps(24) = "吳雅婷" : arrQtrs(24) = "Q1" : arrAmounts(24) = 45000
arrEmps(25) = "吳雅婷" : arrQtrs(25) = "Q2" : arrAmounts(25) = 52000
arrEmps(26) = "吳雅婷" : arrQtrs(26) = "Q3" : arrAmounts(26) = 48000
arrEmps(27) = "吳雅婷" : arrQtrs(27) = "Q4" : arrAmounts(27) = 67000

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
objDataSheet.Cells(1, 1).Value = "業務員"
objDataSheet.Cells(1, 2).Value = "季度"
objDataSheet.Cells(1, 3).Value = "銷售額"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 27
    objDataSheet.Cells(i + 2, 1).Value = arrEmps(i)
    objDataSheet.Cells(i + 2, 2).Value = arrQtrs(i)
    objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
Next

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C29"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、值欄位 ──────────────────────────────────────────
Set objField = objPivot.PivotFields("業務員")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 銷售額"

' ── 設定前 N 名篩選（依銷售額加總降冪取前 5）──────────────────
' AutoShow(Type, Range, Count, Field)
' xlAutomatic=-4105, xlTop=1, Count=5, Field=值欄位名稱
objPivot.PivotFields("業務員").AutoShow xlAutomatic, xlTop, TOP_N, "加總 - 銷售額"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "前 N 名篩選樞紐分析表：銷售額前 " & TOP_N & " 名業務員"
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
