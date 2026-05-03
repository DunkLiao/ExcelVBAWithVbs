' ============================================================
' PivotWithCalculatedField.vbs
' 說明：使用 VBScript 自動建立含自訂計算欄位的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「產品收支」工作表填入收入與成本示範資料
'   3. 建立樞紐分析表，並新增計算欄位「毛利 = 收入 - 成本」
'   4. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithCalculatedField.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "產品收支"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "計算欄位樞紐"
Const OUTPUT_FILE = "05_PivotWithCalculatedField.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157

' ── 範例資料（地區、產品、收入、成本）──────────────────────
Dim arrRegions(15)
Dim arrProducts(15)
Dim arrRevenues(15)
Dim arrCosts(15)

arrRegions(0)  = "北區" : arrProducts(0)  = "筆電" : arrRevenues(0)  = 120000 : arrCosts(0)  = 85000
arrRegions(1)  = "北區" : arrProducts(1)  = "平板" : arrRevenues(1)  = 75000  : arrCosts(1)  = 48000
arrRegions(2)  = "北區" : arrProducts(2)  = "手機" : arrRevenues(2)  = 90000  : arrCosts(2)  = 61000
arrRegions(3)  = "北區" : arrProducts(3)  = "筆電" : arrRevenues(3)  = 135000 : arrCosts(3)  = 92000
arrRegions(4)  = "南區" : arrProducts(4)  = "筆電" : arrRevenues(4)  = 110000 : arrCosts(4)  = 78000
arrRegions(5)  = "南區" : arrProducts(5)  = "平板" : arrRevenues(5)  = 68000  : arrCosts(5)  = 44000
arrRegions(6)  = "南區" : arrProducts(6)  = "手機" : arrRevenues(6)  = 83000  : arrCosts(6)  = 56000
arrRegions(7)  = "南區" : arrProducts(7)  = "平板" : arrRevenues(7)  = 79000  : arrCosts(7)  = 51000
arrRegions(8)  = "東區" : arrProducts(8)  = "筆電" : arrRevenues(8)  = 145000 : arrCosts(8)  = 98000
arrRegions(9)  = "東區" : arrProducts(9)  = "手機" : arrRevenues(9)  = 97000  : arrCosts(9)  = 65000
arrRegions(10) = "東區" : arrProducts(10) = "平板" : arrRevenues(10) = 62000  : arrCosts(10) = 40000
arrRegions(11) = "東區" : arrProducts(11) = "手機" : arrRevenues(11) = 88000  : arrCosts(11) = 59000
arrRegions(12) = "西區" : arrProducts(12) = "筆電" : arrRevenues(12) = 105000 : arrCosts(12) = 73000
arrRegions(13) = "西區" : arrProducts(13) = "手機" : arrRevenues(13) = 76000  : arrCosts(13) = 52000
arrRegions(14) = "西區" : arrProducts(14) = "平板" : arrRevenues(14) = 54000  : arrCosts(14) = 36000
arrRegions(15) = "西區" : arrProducts(15) = "筆電" : arrRevenues(15) = 118000 : arrCosts(15) = 81000

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
objDataSheet.Cells(1, 1).Value = "地區"
objDataSheet.Cells(1, 2).Value = "產品"
objDataSheet.Cells(1, 3).Value = "收入"
objDataSheet.Cells(1, 4).Value = "成本"

With objDataSheet.Range("A1:D1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 15
    objDataSheet.Cells(i + 2, 1).Value = arrRegions(i)
    objDataSheet.Cells(i + 2, 2).Value = arrProducts(i)
    objDataSheet.Cells(i + 2, 3).Value = arrRevenues(i)
    objDataSheet.Cells(i + 2, 4).Value = arrCosts(i)
Next

objDataSheet.Columns("A:D").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:D17"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 新增計算欄位「毛利 = 收入 - 成本」──────────────────────
objPivot.CalculatedFields.Add "毛利", "= 收入 - 成本"

' ── 設定列、欄、值欄位 ──────────────────────────────────────
Set objField = objPivot.PivotFields("地區")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("產品")
objField.Orientation = xlColumnField
objField.Position    = 1

Set objField = objPivot.PivotFields("收入")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 收入"

Set objField = objPivot.PivotFields("成本")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 成本"

Set objField = objPivot.PivotFields("毛利")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "計算 - 毛利"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "含計算欄位的樞紐分析表：毛利 = 收入 - 成本"
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
