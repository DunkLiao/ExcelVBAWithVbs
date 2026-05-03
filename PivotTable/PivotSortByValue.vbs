' ============================================================
' PivotSortByValue.vbs
' 說明：使用 VBScript 自動建立依值欄位排序的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「電商訂單」工作表填入電商平台銷售示範資料
'   3. 建立樞紐分析表（列=商品類別，值=訂單金額加總）
'   4. 將列欄位依訂單金額加總「降冪」自動排序
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotSortByValue.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "電商訂單"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "排序樞紐"
Const OUTPUT_FILE = "11_PivotSortByValue.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlDataField   = 3
Const xlSum         = -4157
Const xlDescending  = 2     ' 降冪排序

' ── 範例資料（平台、商品類別、訂單金額）────────────────────
Dim arrPlatforms(23)
Dim arrCategories(23)
Dim arrAmounts(23)

arrPlatforms(0)  = "蝦皮" : arrCategories(0)  = "3C 電子" : arrAmounts(0)  = 152000
arrPlatforms(1)  = "蝦皮" : arrCategories(1)  = "服飾配件" : arrAmounts(1)  = 87000
arrPlatforms(2)  = "蝦皮" : arrCategories(2)  = "美妝保養" : arrAmounts(2)  = 63000
arrPlatforms(3)  = "蝦皮" : arrCategories(3)  = "食品飲料" : arrAmounts(3)  = 41000
arrPlatforms(4)  = "蝦皮" : arrCategories(4)  = "居家生活" : arrAmounts(4)  = 55000
arrPlatforms(5)  = "蝦皮" : arrCategories(5)  = "運動戶外" : arrAmounts(5)  = 38000
arrPlatforms(6)  = "momo" : arrCategories(6)  = "3C 電子" : arrAmounts(6)  = 198000
arrPlatforms(7)  = "momo" : arrCategories(7)  = "服飾配件" : arrAmounts(7)  = 72000
arrPlatforms(8)  = "momo" : arrCategories(8)  = "美妝保養" : arrAmounts(8)  = 95000
arrPlatforms(9)  = "momo" : arrCategories(9)  = "食品飲料" : arrAmounts(9)  = 118000
arrPlatforms(10) = "momo" : arrCategories(10) = "居家生活" : arrAmounts(10) = 67000
arrPlatforms(11) = "momo" : arrCategories(11) = "運動戶外" : arrAmounts(11) = 44000
arrPlatforms(12) = "PChome" : arrCategories(12) = "3C 電子" : arrAmounts(12) = 231000
arrPlatforms(13) = "PChome" : arrCategories(13) = "服飾配件" : arrAmounts(13) = 48000
arrPlatforms(14) = "PChome" : arrCategories(14) = "美妝保養" : arrAmounts(14) = 39000
arrPlatforms(15) = "PChome" : arrCategories(15) = "食品飲料" : arrAmounts(15) = 76000
arrPlatforms(16) = "PChome" : arrCategories(16) = "居家生活" : arrAmounts(16) = 52000
arrPlatforms(17) = "PChome" : arrCategories(17) = "運動戶外" : arrAmounts(17) = 31000
arrPlatforms(18) = "Yahoo" : arrCategories(18) = "3C 電子" : arrAmounts(18) = 88000
arrPlatforms(19) = "Yahoo" : arrCategories(19) = "服飾配件" : arrAmounts(19) = 61000
arrPlatforms(20) = "Yahoo" : arrCategories(20) = "美妝保養" : arrAmounts(20) = 54000
arrPlatforms(21) = "Yahoo" : arrCategories(21) = "食品飲料" : arrAmounts(21) = 83000
arrPlatforms(22) = "Yahoo" : arrCategories(22) = "居家生活" : arrAmounts(22) = 47000
arrPlatforms(23) = "Yahoo" : arrCategories(23) = "運動戶外" : arrAmounts(23) = 29000

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
objDataSheet.Cells(1, 1).Value = "平台"
objDataSheet.Cells(1, 2).Value = "商品類別"
objDataSheet.Cells(1, 3).Value = "訂單金額"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 23
    objDataSheet.Cells(i + 2, 1).Value = arrPlatforms(i)
    objDataSheet.Cells(i + 2, 2).Value = arrCategories(i)
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

' ── 設定列、值欄位 ──────────────────────────────────────────
Set objField = objPivot.PivotFields("商品類別")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("訂單金額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 訂單金額"

' ── 依值欄位降冪排序列項目 ──────────────────────────────────
' AutoSort(Order, Field)：xlDescending=2，依「加總 - 訂單金額」排序
objPivot.PivotFields("商品類別").AutoSort xlDescending, "加總 - 訂單金額"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "依值排序樞紐分析表：商品類別依訂單金額加總降冪排列"
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
