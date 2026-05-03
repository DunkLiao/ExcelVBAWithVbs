' ============================================================
' PivotWithRunningTotal.vbs
' 說明：使用 VBScript 自動建立以累計加總顯示數值的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「門市銷售」工作表填入逐月銷售示範資料
'   3. 建立樞紐分析表（列=月份，欄=門市，值=銷售額）
'   4. 新增第二個值欄位以「逐月累計加總」方式顯示
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithRunningTotal.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "門市銷售"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "累計樞紐"
Const OUTPUT_FILE = "12_PivotWithRunningTotal.xlsx"

Const xlDatabase     = 1
Const xlRowField     = 1
Const xlColumnField  = 2
Const xlDataField    = 3
Const xlSum          = -4157
Const xlRunningTotal = 5    ' 累計加總

' ── 範例資料（月份、門市、銷售額）──────────────────────────
Dim arrMonths(23)
Dim arrStores(23)
Dim arrAmounts(23)

arrMonths(0)  = 1  : arrStores(0)  = "信義店" : arrAmounts(0)  = 82000
arrMonths(1)  = 1  : arrStores(1)  = "西門店" : arrAmounts(1)  = 65000
arrMonths(2)  = 2  : arrStores(2)  = "信義店" : arrAmounts(2)  = 74000
arrMonths(3)  = 2  : arrStores(3)  = "西門店" : arrAmounts(3)  = 58000
arrMonths(4)  = 3  : arrStores(4)  = "信義店" : arrAmounts(4)  = 91000
arrMonths(5)  = 3  : arrStores(5)  = "西門店" : arrAmounts(5)  = 72000
arrMonths(6)  = 4  : arrStores(6)  = "信義店" : arrAmounts(6)  = 105000
arrMonths(7)  = 4  : arrStores(7)  = "西門店" : arrAmounts(7)  = 83000
arrMonths(8)  = 5  : arrStores(8)  = "信義店" : arrAmounts(8)  = 118000
arrMonths(9)  = 5  : arrStores(9)  = "西門店" : arrAmounts(9)  = 94000
arrMonths(10) = 6  : arrStores(10) = "信義店" : arrAmounts(10) = 132000
arrMonths(11) = 6  : arrStores(11) = "西門店" : arrAmounts(11) = 105000
arrMonths(12) = 7  : arrStores(12) = "信義店" : arrAmounts(12) = 125000
arrMonths(13) = 7  : arrStores(13) = "西門店" : arrAmounts(13) = 98000
arrMonths(14) = 8  : arrStores(14) = "信義店" : arrAmounts(14) = 141000
arrMonths(15) = 8  : arrStores(15) = "西門店" : arrAmounts(15) = 112000
arrMonths(16) = 9  : arrStores(16) = "信義店" : arrAmounts(16) = 158000
arrMonths(17) = 9  : arrStores(17) = "西門店" : arrAmounts(17) = 126000
arrMonths(18) = 10 : arrStores(18) = "信義店" : arrAmounts(18) = 175000
arrMonths(19) = 10 : arrStores(19) = "西門店" : arrAmounts(19) = 138000
arrMonths(20) = 11 : arrStores(20) = "信義店" : arrAmounts(20) = 210000
arrMonths(21) = 11 : arrStores(21) = "西門店" : arrAmounts(21) = 168000
arrMonths(22) = 12 : arrStores(22) = "信義店" : arrAmounts(22) = 248000
arrMonths(23) = 12 : arrStores(23) = "西門店" : arrAmounts(23) = 195000

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
objDataSheet.Cells(1, 1).Value = "月份"
objDataSheet.Cells(1, 2).Value = "門市"
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
    objDataSheet.Cells(i + 2, 2).Value = arrStores(i)
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

Set objField = objPivot.PivotFields("門市")
objField.Orientation = xlColumnField
objField.Position    = 1

' 第一個值欄位：一般加總
Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "當月銷售額"

' 第二個值欄位：累計加總
Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "累計銷售額"

' ── 設定累計計算方式（沿月份欄位累計）──────────────────────
Set objDataField = objPivot.DataFields("累計銷售額")
objDataField.Calculation = xlRunningTotal
objDataField.BaseField   = "月份"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "累計加總樞紐分析表：同時顯示當月銷售額與逐月累計金額"
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
