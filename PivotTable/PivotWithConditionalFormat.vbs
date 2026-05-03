' ============================================================
' PivotWithConditionalFormat.vbs
' 說明：使用 VBScript 自動建立套用條件格式化的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「庫存資料」工作表填入倉庫庫存示範資料
'   3. 建立樞紐分析表
'   4. 在樞紐分析表值區域套用三色色階條件格式化（紅→黃→綠）
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithConditionalFormat.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "庫存資料"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "條件格式樞紐"
Const OUTPUT_FILE = "08_PivotWithConditionalFormat.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157

' 色階條件格式化類型常數
Const xlConditionValueLowestValue  = 1
Const xlConditionValueHighestValue = 2
Const xlConditionValuePercentile   = 5

' ── 範例資料（倉庫、商品類別、庫存量）──────────────────────
Dim arrWarehouses(19)
Dim arrCategories(19)
Dim arrStocks(19)

arrWarehouses(0)  = "台北倉" : arrCategories(0)  = "電子產品" : arrStocks(0)  = 520
arrWarehouses(1)  = "台北倉" : arrCategories(1)  = "服飾用品" : arrStocks(1)  = 310
arrWarehouses(2)  = "台北倉" : arrCategories(2)  = "食品飲料" : arrStocks(2)  = 850
arrWarehouses(3)  = "台北倉" : arrCategories(3)  = "家居用品" : arrStocks(3)  = 230
arrWarehouses(4)  = "台北倉" : arrCategories(4)  = "運動器材" : arrStocks(4)  = 180
arrWarehouses(5)  = "台中倉" : arrCategories(5)  = "電子產品" : arrStocks(5)  = 380
arrWarehouses(6)  = "台中倉" : arrCategories(6)  = "服飾用品" : arrStocks(6)  = 420
arrWarehouses(7)  = "台中倉" : arrCategories(7)  = "食品飲料" : arrStocks(7)  = 640
arrWarehouses(8)  = "台中倉" : arrCategories(8)  = "家居用品" : arrStocks(8)  = 290
arrWarehouses(9)  = "台中倉" : arrCategories(9)  = "運動器材" : arrStocks(9)  = 150
arrWarehouses(10) = "高雄倉" : arrCategories(10) = "電子產品" : arrStocks(10) = 610
arrWarehouses(11) = "高雄倉" : arrCategories(11) = "服飾用品" : arrStocks(11) = 270
arrWarehouses(12) = "高雄倉" : arrCategories(12) = "食品飲料" : arrStocks(12) = 920
arrWarehouses(13) = "高雄倉" : arrCategories(13) = "家居用品" : arrStocks(13) = 340
arrWarehouses(14) = "高雄倉" : arrCategories(14) = "運動器材" : arrStocks(14) = 210
arrWarehouses(15) = "桃園倉" : arrCategories(15) = "電子產品" : arrStocks(15) = 450
arrWarehouses(16) = "桃園倉" : arrCategories(16) = "服飾用品" : arrStocks(16) = 380
arrWarehouses(17) = "桃園倉" : arrCategories(17) = "食品飲料" : arrStocks(17) = 710
arrWarehouses(18) = "桃園倉" : arrCategories(18) = "家居用品" : arrStocks(18) = 195
arrWarehouses(19) = "桃園倉" : arrCategories(19) = "運動器材" : arrStocks(19) = 125

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objDataSheet, objPivotSheet
Dim objCache, objPivot, objField, objColorScale
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
objDataSheet.Cells(1, 1).Value = "倉庫"
objDataSheet.Cells(1, 2).Value = "商品類別"
objDataSheet.Cells(1, 3).Value = "庫存量"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 19
    objDataSheet.Cells(i + 2, 1).Value = arrWarehouses(i)
    objDataSheet.Cells(i + 2, 2).Value = arrCategories(i)
    objDataSheet.Cells(i + 2, 3).Value = arrStocks(i)
Next

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C21"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、欄、值欄位 ──────────────────────────────────────
Set objField = objPivot.PivotFields("倉庫")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("商品類別")
objField.Orientation = xlColumnField
objField.Position    = 1

Set objField = objPivot.PivotFields("庫存量")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 庫存量"

' ── 在值區域套用三色色階（紅低→黃中→綠高）─────────────────
Set objColorScale = objPivot.DataBodyRange.FormatConditions.AddColorScale(3)

With objColorScale.ColorScaleCriteria(1)
    .Type = xlConditionValueLowestValue
    .FormatColor.Color = RGB(248, 105, 107)  ' 紅色：最低值
End With

With objColorScale.ColorScaleCriteria(2)
    .Type = xlConditionValuePercentile
    .Value = 50
    .FormatColor.Color = RGB(255, 235, 132)  ' 黃色：中間值（第 50 百分位）
End With

With objColorScale.ColorScaleCriteria(3)
    .Type = xlConditionValueHighestValue
    .FormatColor.Color = RGB(99, 190, 123)   ' 綠色：最高值
End With

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "條件格式化樞紐分析表：庫存量色階（紅=低 / 黃=中 / 綠=高）"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 14
End With

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objColorScale = Nothing
Set objField      = Nothing
Set objPivot      = Nothing
Set objCache      = Nothing
Set objPivotSheet = Nothing
Set objDataSheet  = Nothing
Set objWorkbook   = Nothing
Set objExcel      = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
