' ============================================================
' BasicPivotTable.vbs
' 說明：使用 VBScript 自動建立 Excel 基本樞紐分析表範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「銷售資料」工作表填入示範銷售資料
'   3. 建立樞紐分析表（列=地區，欄=產品，值=銷售額加總）
'   4. 格式化樞紐分析表標題
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\BasicPivotTable.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA   = "銷售資料"
Const SHEET_PIVOT  = "樞紐分析表"
Const PIVOT_NAME   = "基本樞紐"
Const OUTPUT_FILE  = "01_BasicPivotTable.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157

' ── 範例資料（地區、產品、銷售額）────────────────────────────
Dim arrData(20, 2)
arrData(0,  0) = "地區" : arrData(0,  1) = "產品" : arrData(0,  2) = "銷售額"
arrData(1,  0) = "北區" : arrData(1,  1) = "筆電" : arrData(1,  2) = 85000
arrData(2,  0) = "北區" : arrData(2,  1) = "平板" : arrData(2,  2) = 52000
arrData(3,  0) = "北區" : arrData(3,  1) = "手機" : arrData(3,  2) = 67000
arrData(4,  0) = "北區" : arrData(4,  1) = "筆電" : arrData(4,  2) = 91000
arrData(5,  0) = "北區" : arrData(5,  1) = "手機" : arrData(5,  2) = 73000
arrData(6,  0) = "南區" : arrData(6,  1) = "筆電" : arrData(6,  2) = 76000
arrData(7,  0) = "南區" : arrData(7,  1) = "平板" : arrData(7,  2) = 48000
arrData(8,  0) = "南區" : arrData(8,  1) = "手機" : arrData(8,  2) = 61000
arrData(9,  0) = "南區" : arrData(9,  1) = "平板" : arrData(9,  2) = 55000
arrData(10, 0) = "南區" : arrData(10, 1) = "筆電" : arrData(10, 2) = 82000
arrData(11, 0) = "東區" : arrData(11, 1) = "手機" : arrData(11, 2) = 79000
arrData(12, 0) = "東區" : arrData(12, 1) = "筆電" : arrData(12, 2) = 93000
arrData(13, 0) = "東區" : arrData(13, 1) = "平板" : arrData(13, 2) = 44000
arrData(14, 0) = "東區" : arrData(14, 1) = "手機" : arrData(14, 2) = 68000
arrData(15, 0) = "東區" : arrData(15, 1) = "平板" : arrData(15, 2) = 50000
arrData(16, 0) = "西區" : arrData(16, 1) = "筆電" : arrData(16, 2) = 71000
arrData(17, 0) = "西區" : arrData(17, 1) = "手機" : arrData(17, 2) = 58000
arrData(18, 0) = "西區" : arrData(18, 1) = "平板" : arrData(18, 2) = 39000
arrData(19, 0) = "西區" : arrData(19, 1) = "筆電" : arrData(19, 2) = 88000
arrData(20, 0) = "西區" : arrData(20, 1) = "手機" : arrData(20, 2) = 62000

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objDataSheet, objPivotSheet
Dim objCache, objPivot, objField
Dim savePath, objShell, r, c

Set objShell = CreateObject("WScript.Shell")
savePath = objShell.SpecialFolders("Desktop") & "\" & OUTPUT_FILE
Set objShell = Nothing

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible       = False
objExcel.DisplayAlerts = False

Set objWorkbook   = objExcel.Workbooks.Add()
Set objDataSheet  = objWorkbook.Sheets(1)
objDataSheet.Name = SHEET_DATA

' ── 寫入示範資料 ────────────────────────────────────────────
For r = 0 To 20
    For c = 0 To 2
        objDataSheet.Cells(r + 1, c + 1).Value = arrData(r, c)
    Next
Next

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C21"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、欄、值欄位 ──────────────────────────────────────
Set objField = objPivot.PivotFields("地區")
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
objPivotSheet.Range("A1").Value = "基本樞紐分析表：各地區產品銷售額加總"
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
