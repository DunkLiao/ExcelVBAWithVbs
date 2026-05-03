' ============================================================
' PivotWithShowPages.vbs
' 說明：使用 VBScript 自動建立並透過 ShowPages 展開工作表的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「分區業績」工作表填入銷售示範資料
'   3. 建立含報表篩選欄位（地區）的樞紐分析表
'   4. 呼叫 ShowPages 自動為每個地區各建立一個獨立工作表
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithShowPages.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "分區業績"
Const SHEET_PIVOT = "總覽樞紐"
Const PIVOT_NAME  = "分頁展開樞紐"
Const OUTPUT_FILE = "17_PivotWithShowPages.xlsx"

Const xlDatabase    = 1
Const xlPageField   = 3
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157

' ── 範例資料（地區、季度、產品、銷售額）────────────────────
Dim arrRegions(23)
Dim arrQtrs(23)
Dim arrProducts(23)
Dim arrAmounts(23)

arrRegions(0)  = "北區" : arrQtrs(0)  = "Q1" : arrProducts(0)  = "A產品" : arrAmounts(0)  = 85000
arrRegions(1)  = "北區" : arrQtrs(1)  = "Q1" : arrProducts(1)  = "B產品" : arrAmounts(1)  = 62000
arrRegions(2)  = "北區" : arrQtrs(2)  = "Q2" : arrProducts(2)  = "A產品" : arrAmounts(2)  = 97000
arrRegions(3)  = "北區" : arrQtrs(3)  = "Q2" : arrProducts(3)  = "B產品" : arrAmounts(3)  = 74000
arrRegions(4)  = "北區" : arrQtrs(4)  = "Q3" : arrProducts(4)  = "A產品" : arrAmounts(4)  = 112000
arrRegions(5)  = "北區" : arrQtrs(5)  = "Q3" : arrProducts(5)  = "B產品" : arrAmounts(5)  = 88000
arrRegions(6)  = "南區" : arrQtrs(6)  = "Q1" : arrProducts(6)  = "A產品" : arrAmounts(6)  = 73000
arrRegions(7)  = "南區" : arrQtrs(7)  = "Q1" : arrProducts(7)  = "B產品" : arrAmounts(7)  = 51000
arrRegions(8)  = "南區" : arrQtrs(8)  = "Q2" : arrProducts(8)  = "A產品" : arrAmounts(8)  = 84000
arrRegions(9)  = "南區" : arrQtrs(9)  = "Q2" : arrProducts(9)  = "B產品" : arrAmounts(9)  = 63000
arrRegions(10) = "南區" : arrQtrs(10) = "Q3" : arrProducts(10) = "A產品" : arrAmounts(10) = 98000
arrRegions(11) = "南區" : arrQtrs(11) = "Q3" : arrProducts(11) = "B產品" : arrAmounts(11) = 75000
arrRegions(12) = "東區" : arrQtrs(12) = "Q1" : arrProducts(12) = "A產品" : arrAmounts(12) = 58000
arrRegions(13) = "東區" : arrQtrs(13) = "Q1" : arrProducts(13) = "B產品" : arrAmounts(13) = 41000
arrRegions(14) = "東區" : arrQtrs(14) = "Q2" : arrProducts(14) = "A產品" : arrAmounts(14) = 67000
arrRegions(15) = "東區" : arrQtrs(15) = "Q2" : arrProducts(15) = "B產品" : arrAmounts(15) = 49000
arrRegions(16) = "東區" : arrQtrs(16) = "Q3" : arrProducts(16) = "A產品" : arrAmounts(16) = 79000
arrRegions(17) = "東區" : arrQtrs(17) = "Q3" : arrProducts(17) = "B產品" : arrAmounts(17) = 58000
arrRegions(18) = "西區" : arrQtrs(18) = "Q1" : arrProducts(18) = "A產品" : arrAmounts(18) = 64000
arrRegions(19) = "西區" : arrQtrs(19) = "Q1" : arrProducts(19) = "B產品" : arrAmounts(19) = 47000
arrRegions(20) = "西區" : arrQtrs(20) = "Q2" : arrProducts(20) = "A產品" : arrAmounts(20) = 75000
arrRegions(21) = "西區" : arrQtrs(21) = "Q2" : arrProducts(21) = "B產品" : arrAmounts(21) = 55000
arrRegions(22) = "西區" : arrQtrs(22) = "Q3" : arrProducts(22) = "A產品" : arrAmounts(22) = 88000
arrRegions(23) = "西區" : arrQtrs(23) = "Q3" : arrProducts(23) = "B產品" : arrAmounts(23) = 64000

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
objDataSheet.Cells(1, 2).Value = "季度"
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
    objDataSheet.Cells(i + 2, 1).Value = arrRegions(i)
    objDataSheet.Cells(i + 2, 2).Value = arrQtrs(i)
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
Set objField = objPivot.PivotFields("地區")
objField.Orientation = xlPageField
objField.Position    = 1

Set objField = objPivot.PivotFields("季度")
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
objPivotSheet.Range("A1").Value = "ShowPages 樞紐分析表：點下方呼叫 ShowPages 可為每地區建立獨立工作表"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 13
End With

' ── 呼叫 ShowPages 為每個地區自動建立獨立工作表 ──────────────
' ShowPages(PageField)：依報表篩選欄位「地區」展開，
' 每個篩選值（北區/南區/東區/西區）各建立一個工作表
objPivot.ShowPages "地區"

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
WScript.Echo "提示：活頁簿中包含「北區」「南區」「東區」「西區」四個由 ShowPages 建立的分區工作表"
