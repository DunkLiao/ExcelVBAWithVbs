' ============================================================
' PivotWithSlicer.vbs
' 說明：使用 VBScript 自動建立含交叉分析篩選器的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「人事薪資」工作表填入員工薪資示範資料
'   3. 建立樞紐分析表（列=部門，值=薪資加總）
'   4. 插入「部門」與「職級」兩個交叉分析篩選器
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithSlicer.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "人事薪資"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "薪資樞紐"
Const OUTPUT_FILE = "09_PivotWithSlicer.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157

' ── 範例資料（部門、職級、員工、薪資）──────────────────────
Dim arrDepts(19)
Dim arrGrades(19)
Dim arrEmps(19)
Dim arrSalaries(19)

arrDepts(0)  = "研發部" : arrGrades(0)  = "資深" : arrEmps(0)  = "王小明" : arrSalaries(0)  = 85000
arrDepts(1)  = "研發部" : arrGrades(1)  = "資深" : arrEmps(1)  = "李大華" : arrSalaries(1)  = 88000
arrDepts(2)  = "研發部" : arrGrades(2)  = "初級" : arrEmps(2)  = "陳美玲" : arrSalaries(2)  = 52000
arrDepts(3)  = "研發部" : arrGrades(3)  = "初級" : arrEmps(3)  = "張志強" : arrSalaries(3)  = 48000
arrDepts(4)  = "研發部" : arrGrades(4)  = "主管"  : arrEmps(4)  = "林佳慧" : arrSalaries(4)  = 120000
arrDepts(5)  = "業務部" : arrGrades(5)  = "資深" : arrEmps(5)  = "黃文成" : arrSalaries(5)  = 72000
arrDepts(6)  = "業務部" : arrGrades(6)  = "資深" : arrEmps(6)  = "吳雅婷" : arrSalaries(6)  = 68000
arrDepts(7)  = "業務部" : arrGrades(7)  = "初級" : arrEmps(7)  = "劉建宏" : arrSalaries(7)  = 45000
arrDepts(8)  = "業務部" : arrGrades(8)  = "初級" : arrEmps(8)  = "周淑芬" : arrSalaries(8)  = 42000
arrDepts(9)  = "業務部" : arrGrades(9)  = "主管"  : arrEmps(9)  = "鄭國強" : arrSalaries(9)  = 110000
arrDepts(10) = "行政部" : arrGrades(10) = "資深" : arrEmps(10) = "謝麗華" : arrSalaries(10) = 58000
arrDepts(11) = "行政部" : arrGrades(11) = "初級" : arrEmps(11) = "許志豪" : arrSalaries(11) = 38000
arrDepts(12) = "行政部" : arrGrades(12) = "初級" : arrEmps(12) = "楊淑惠" : arrSalaries(12) = 36000
arrDepts(13) = "行政部" : arrGrades(13) = "主管"  : arrEmps(13) = "蔡明宏" : arrSalaries(13) = 90000
arrDepts(14) = "財務部" : arrGrades(14) = "資深" : arrEmps(14) = "洪雅君" : arrSalaries(14) = 78000
arrDepts(15) = "財務部" : arrGrades(15) = "資深" : arrEmps(15) = "林冠廷" : arrSalaries(15) = 75000
arrDepts(16) = "財務部" : arrGrades(16) = "初級" : arrEmps(16) = "賴怡婷" : arrSalaries(16) = 50000
arrDepts(17) = "財務部" : arrGrades(17) = "初級" : arrEmps(17) = "葉俊男" : arrSalaries(17) = 47000
arrDepts(18) = "財務部" : arrGrades(18) = "主管"  : arrEmps(18) = "吳明哲" : arrSalaries(18) = 105000
arrDepts(19) = "研發部" : arrGrades(19) = "資深" : arrEmps(19) = "方思穎" : arrSalaries(19) = 82000

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objDataSheet, objPivotSheet
Dim objCache, objPivot, objField
Dim objSlicerCache1, objSlicerCache2
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
objDataSheet.Cells(1, 2).Value = "職級"
objDataSheet.Cells(1, 3).Value = "員工"
objDataSheet.Cells(1, 4).Value = "薪資"

With objDataSheet.Range("A1:D1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 19
    objDataSheet.Cells(i + 2, 1).Value = arrDepts(i)
    objDataSheet.Cells(i + 2, 2).Value = arrGrades(i)
    objDataSheet.Cells(i + 2, 3).Value = arrEmps(i)
    objDataSheet.Cells(i + 2, 4).Value = arrSalaries(i)
Next

objDataSheet.Columns("A:D").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:D21"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、欄、值欄位 ──────────────────────────────────────
Set objField = objPivot.PivotFields("部門")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("職級")
objField.Orientation = xlColumnField
objField.Position    = 1

Set objField = objPivot.PivotFields("薪資")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 薪資"

' ── 插入「部門」交叉分析篩選器 ──────────────────────────────
' Add2(Source, SourceField, [Name])
Set objSlicerCache1 = objWorkbook.SlicerCaches.Add2(objPivot, "部門")
' Slicers.Add(SlicerDestination, [Level], [Name], [Caption], [Top], [Left], [Width], [Height])
objSlicerCache1.Slicers.Add objPivotSheet, , "部門篩選器", "部門", 20, 380, 160, 200

' ── 插入「職級」交叉分析篩選器 ──────────────────────────────
Set objSlicerCache2 = objWorkbook.SlicerCaches.Add2(objPivot, "職級")
objSlicerCache2.Slicers.Add objPivotSheet, , "職級篩選器", "職級", 240, 380, 160, 160

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "含交叉分析篩選器的樞紐分析表：可依部門與職級互動篩選薪資"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 14
End With

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objSlicerCache2 = Nothing
Set objSlicerCache1 = Nothing
Set objField        = Nothing
Set objPivot        = Nothing
Set objCache        = Nothing
Set objPivotSheet   = Nothing
Set objDataSheet    = Nothing
Set objWorkbook     = Nothing
Set objExcel        = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
