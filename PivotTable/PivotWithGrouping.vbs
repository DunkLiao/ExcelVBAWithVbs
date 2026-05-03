' ============================================================
' PivotWithGrouping.vbs
' 說明：使用 VBScript 自動建立含數值群組功能的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「月份業績」工作表填入 12 個月份的銷售示範資料
'   3. 建立樞紐分析表後，將月份欄位（數值 1-12）群組為每季（步距 3）
'   4. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithGrouping.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "月份業績"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "月份群組樞紐"
Const OUTPUT_FILE = "03_PivotWithGrouping.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157

' ── 範例資料（月份數字 1-12、通路、銷售額）──────────────────
' 24 筆：每月 × 2 通路（直營/代理）
Dim arrMonths(23)
Dim arrChannels(23)
Dim arrAmounts(23)

arrMonths(0)  = 1  : arrChannels(0)  = "直營" : arrAmounts(0)  = 65000
arrMonths(1)  = 1  : arrChannels(1)  = "代理" : arrAmounts(1)  = 48000
arrMonths(2)  = 2  : arrChannels(2)  = "直營" : arrAmounts(2)  = 58000
arrMonths(3)  = 2  : arrChannels(3)  = "代理" : arrAmounts(3)  = 41000
arrMonths(4)  = 3  : arrChannels(4)  = "直營" : arrAmounts(4)  = 72000
arrMonths(5)  = 3  : arrChannels(5)  = "代理" : arrAmounts(5)  = 55000
arrMonths(6)  = 4  : arrChannels(6)  = "直營" : arrAmounts(6)  = 80000
arrMonths(7)  = 4  : arrChannels(7)  = "代理" : arrAmounts(7)  = 63000
arrMonths(8)  = 5  : arrChannels(8)  = "直營" : arrAmounts(8)  = 91000
arrMonths(9)  = 5  : arrChannels(9)  = "代理" : arrAmounts(9)  = 70000
arrMonths(10) = 6  : arrChannels(10) = "直營" : arrAmounts(10) = 85000
arrMonths(11) = 6  : arrChannels(11) = "代理" : arrAmounts(11) = 67000
arrMonths(12) = 7  : arrChannels(12) = "直營" : arrAmounts(12) = 94000
arrMonths(13) = 7  : arrChannels(13) = "代理" : arrAmounts(13) = 75000
arrMonths(14) = 8  : arrChannels(14) = "直營" : arrAmounts(14) = 88000
arrMonths(15) = 8  : arrChannels(15) = "代理" : arrAmounts(15) = 69000
arrMonths(16) = 9  : arrChannels(16) = "直營" : arrAmounts(16) = 102000
arrMonths(17) = 9  : arrChannels(17) = "代理" : arrAmounts(17) = 81000
arrMonths(18) = 10 : arrChannels(18) = "直營" : arrAmounts(18) = 110000
arrMonths(19) = 10 : arrChannels(19) = "代理" : arrAmounts(19) = 88000
arrMonths(20) = 11 : arrChannels(20) = "直營" : arrAmounts(20) = 125000
arrMonths(21) = 11 : arrChannels(21) = "代理" : arrAmounts(21) = 97000
arrMonths(22) = 12 : arrChannels(22) = "直營" : arrAmounts(22) = 138000
arrMonths(23) = 12 : arrChannels(23) = "代理" : arrAmounts(23) = 109000

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
objDataSheet.Cells(1, 1).Value = "月份"
objDataSheet.Cells(1, 2).Value = "通路"
objDataSheet.Cells(1, 3).Value = "銷售額"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 23
    objDataSheet.Cells(i + 2, 1).Value = arrMonths(i)
    objDataSheet.Cells(i + 2, 2).Value = arrChannels(i)
    objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
Next

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C25"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、欄、值欄位 ──────────────────────────────────────
Set objField = objPivot.PivotFields("月份")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("通路")
objField.Orientation = xlColumnField
objField.Position    = 1

Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 銷售額"

' ── 將月份欄位依步距 3 群組（1-3, 4-6, 7-9, 10-12）────────────
' Group(Start, End, By)：從 1 到 12，步距 3
objPivot.PivotFields("月份").Group 1, 12, 3

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "含群組功能的樞紐分析表：月份自動群組為每季（Q1-Q4）"
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
