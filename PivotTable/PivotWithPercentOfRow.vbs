' ============================================================
' PivotWithPercentOfRow.vbs
' 說明：使用 VBScript 自動建立以列百分比顯示的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「廣告投放」工作表填入行銷費用示範資料
'   3. 建立樞紐分析表（列=行銷管道，欄=產品，值=費用加總）
'   4. 新增第二個值欄位，以「列百分比」方式顯示
'      （各欄佔同列總計的百分比，分析各管道費用分配）
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithPercentOfRow.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "廣告投放"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "列百分比樞紐"
Const OUTPUT_FILE = "20_PivotWithPercentOfRow.xlsx"

Const xlDatabase      = 1
Const xlRowField      = 1
Const xlColumnField   = 2
Const xlDataField     = 3
Const xlSum           = -4157
Const xlPercentOfRow  = 6    ' 列百分比（佔同列加總的比例）

' ── 範例資料（行銷管道、產品線、廣告費用）──────────────────
Dim arrChannels(19)
Dim arrProducts(19)
Dim arrBudgets(19)

arrChannels(0)  = "搜尋廣告" : arrProducts(0)  = "旗艦機" : arrBudgets(0)  = 85000
arrChannels(1)  = "搜尋廣告" : arrProducts(1)  = "標準機" : arrBudgets(1)  = 62000
arrChannels(2)  = "搜尋廣告" : arrProducts(2)  = "入門機" : arrBudgets(2)  = 38000
arrChannels(3)  = "搜尋廣告" : arrProducts(3)  = "平板"   : arrBudgets(3)  = 45000
arrChannels(4)  = "社群媒體" : arrProducts(4)  = "旗艦機" : arrBudgets(4)  = 125000
arrChannels(5)  = "社群媒體" : arrProducts(5)  = "標準機" : arrBudgets(5)  = 98000
arrChannels(6)  = "社群媒體" : arrProducts(6)  = "入門機" : arrBudgets(6)  = 54000
arrChannels(7)  = "社群媒體" : arrProducts(7)  = "平板"   : arrBudgets(7)  = 72000
arrChannels(8)  = "電視廣告" : arrProducts(8)  = "旗艦機" : arrBudgets(8)  = 210000
arrChannels(9)  = "電視廣告" : arrProducts(9)  = "標準機" : arrBudgets(9)  = 180000
arrChannels(10) = "電視廣告" : arrProducts(10) = "入門機" : arrBudgets(10) = 95000
arrChannels(11) = "電視廣告" : arrProducts(11) = "平板"   : arrBudgets(11) = 115000
arrChannels(12) = "網紅合作" : arrProducts(12) = "旗艦機" : arrBudgets(12) = 68000
arrChannels(13) = "網紅合作" : arrProducts(13) = "標準機" : arrBudgets(13) = 52000
arrChannels(14) = "網紅合作" : arrProducts(14) = "入門機" : arrBudgets(14) = 31000
arrChannels(15) = "網紅合作" : arrProducts(15) = "平板"   : arrBudgets(15) = 49000
arrChannels(16) = "戶外廣告" : arrProducts(16) = "旗艦機" : arrBudgets(16) = 145000
arrChannels(17) = "戶外廣告" : arrProducts(17) = "標準機" : arrBudgets(17) = 112000
arrChannels(18) = "戶外廣告" : arrProducts(18) = "入門機" : arrBudgets(18) = 67000
arrChannels(19) = "戶外廣告" : arrProducts(19) = "平板"   : arrBudgets(19) = 83000

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
objDataSheet.Cells(1, 1).Value = "行銷管道"
objDataSheet.Cells(1, 2).Value = "產品線"
objDataSheet.Cells(1, 3).Value = "廣告費用"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 19
    objDataSheet.Cells(i + 2, 1).Value = arrChannels(i)
    objDataSheet.Cells(i + 2, 2).Value = arrProducts(i)
    objDataSheet.Cells(i + 2, 3).Value = arrBudgets(i)
Next

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C21"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、欄欄位 ──────────────────────────────────────────
Set objField = objPivot.PivotFields("行銷管道")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("產品線")
objField.Orientation = xlColumnField
objField.Position    = 1

' 第一個值欄位：費用絕對金額
Set objField = objPivot.PivotFields("廣告費用")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "廣告費用（元）"

' 第二個值欄位：列百分比
Set objField = objPivot.PivotFields("廣告費用")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "列佔比（%）"

' ── 設定列百分比計算方式 ────────────────────────────────────
Set objDataField = objPivot.DataFields("列佔比（%）")
objDataField.Calculation  = xlPercentOfRow
objDataField.NumberFormat = "0.0%"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "列百分比樞紐分析表：各管道費用同時顯示金額與佔該管道總費用之比例"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 13
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
