' ============================================================
' PivotWithPercentage.vbs
' 說明：使用 VBScript 自動建立以百分比顯示數值的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「市場銷售」工作表填入銷售示範資料
'   3. 建立樞紐分析表，值欄位改以「總計百分比」方式顯示
'   4. 設定百分比格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithPercentage.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "市場銷售"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "百分比樞紐"
Const OUTPUT_FILE = "06_PivotWithPercentage.xlsx"

Const xlDatabase         = 1
Const xlRowField         = 1
Const xlColumnField      = 2
Const xlDataField        = 3
Const xlSum              = -4157
Const xlPercentOfTotal   = 9   ' 總計百分比

' ── 範例資料（品牌、通路、銷售額）──────────────────────────
Dim arrBrands(15)
Dim arrChannels(15)
Dim arrAmounts(15)

arrBrands(0)  = "品牌A" : arrChannels(0)  = "實體店" : arrAmounts(0)  = 152000
arrBrands(1)  = "品牌A" : arrChannels(1)  = "網路商城" : arrAmounts(1)  = 98000
arrBrands(2)  = "品牌A" : arrChannels(2)  = "代理商" : arrAmounts(2)  = 74000
arrBrands(3)  = "品牌A" : arrChannels(3)  = "實體店" : arrAmounts(3)  = 138000
arrBrands(4)  = "品牌B" : arrChannels(4)  = "實體店" : arrAmounts(4)  = 87000
arrBrands(5)  = "品牌B" : arrChannels(5)  = "網路商城" : arrAmounts(5)  = 124000
arrBrands(6)  = "品牌B" : arrChannels(6)  = "代理商" : arrAmounts(6)  = 63000
arrBrands(7)  = "品牌B" : arrChannels(7)  = "網路商城" : arrAmounts(7)  = 109000
arrBrands(8)  = "品牌C" : arrChannels(8)  = "實體店" : arrAmounts(8)  = 76000
arrBrands(9)  = "品牌C" : arrChannels(9)  = "網路商城" : arrAmounts(9)  = 55000
arrBrands(10) = "品牌C" : arrChannels(10) = "代理商" : arrAmounts(10) = 92000
arrBrands(11) = "品牌C" : arrChannels(11) = "代理商" : arrAmounts(11) = 81000
arrBrands(12) = "品牌D" : arrChannels(12) = "實體店" : arrAmounts(12) = 113000
arrBrands(13) = "品牌D" : arrChannels(13) = "網路商城" : arrAmounts(13) = 67000
arrBrands(14) = "品牌D" : arrChannels(14) = "代理商" : arrAmounts(14) = 49000
arrBrands(15) = "品牌D" : arrChannels(15) = "實體店" : arrAmounts(15) = 95000

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
objDataSheet.Cells(1, 1).Value = "品牌"
objDataSheet.Cells(1, 2).Value = "通路"
objDataSheet.Cells(1, 3).Value = "銷售額"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 15
    objDataSheet.Cells(i + 2, 1).Value = arrBrands(i)
    objDataSheet.Cells(i + 2, 2).Value = arrChannels(i)
    objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
Next

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C17"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、欄、值欄位 ──────────────────────────────────────
Set objField = objPivot.PivotFields("品牌")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("通路")
objField.Orientation = xlColumnField
objField.Position    = 1

Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "佔比 - 銷售額"

' ── 設定值欄位以總計百分比顯示 ──────────────────────────────
Set objDataField = objPivot.DataFields("佔比 - 銷售額")
objDataField.Calculation   = xlPercentOfTotal
objDataField.NumberFormat  = "0.00%"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "百分比樞紐分析表：各品牌各通路佔總體銷售額的比例"
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
