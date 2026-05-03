' ============================================================
' PivotWithMultipleRowFields.vbs
' 說明：使用 VBScript 自動建立含三層巢狀列欄位的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「通路銷售」工作表填入含大區/城市/通路的銷售示範資料
'   3. 建立樞紐分析表（列=大區 > 城市 > 通路，值=銷售額加總）
'   4. 展示三層巢狀列欄位的樞紐分析表結構
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithMultipleRowFields.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "通路銷售"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "巢狀列欄位樞紐"
Const OUTPUT_FILE = "14_PivotWithMultipleRowFields.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlDataField   = 3
Const xlSum         = -4157

' ── 範例資料（大區、城市、通路、銷售額）────────────────────
Dim arrZones(23)
Dim arrCities(23)
Dim arrChannels(23)
Dim arrAmounts(23)

arrZones(0)  = "北部" : arrCities(0)  = "台北" : arrChannels(0)  = "直營門市" : arrAmounts(0)  = 182000
arrZones(1)  = "北部" : arrCities(1)  = "台北" : arrChannels(1)  = "網路商城" : arrAmounts(1)  = 135000
arrZones(2)  = "北部" : arrCities(2)  = "台北" : arrChannels(2)  = "代理商"   : arrAmounts(2)  = 98000
arrZones(3)  = "北部" : arrCities(3)  = "桃園" : arrChannels(3)  = "直營門市" : arrAmounts(3)  = 124000
arrZones(4)  = "北部" : arrCities(4)  = "桃園" : arrChannels(4)  = "網路商城" : arrAmounts(4)  = 87000
arrZones(5)  = "北部" : arrCities(5)  = "新竹" : arrChannels(5)  = "直營門市" : arrAmounts(5)  = 96000
arrZones(6)  = "北部" : arrCities(6)  = "新竹" : arrChannels(6)  = "代理商"   : arrAmounts(6)  = 73000
arrZones(7)  = "中部" : arrCities(7)  = "台中" : arrChannels(7)  = "直營門市" : arrAmounts(7)  = 156000
arrZones(8)  = "中部" : arrCities(8)  = "台中" : arrChannels(8)  = "網路商城" : arrAmounts(8)  = 112000
arrZones(9)  = "中部" : arrCities(9)  = "台中" : arrChannels(9)  = "代理商"   : arrAmounts(9)  = 84000
arrZones(10) = "中部" : arrCities(10) = "彰化" : arrChannels(10) = "直營門市" : arrAmounts(10) = 78000
arrZones(11) = "中部" : arrCities(11) = "彰化" : arrChannels(11) = "代理商"   : arrAmounts(11) = 52000
arrZones(12) = "南部" : arrCities(12) = "高雄" : arrChannels(12) = "直營門市" : arrAmounts(12) = 168000
arrZones(13) = "南部" : arrCities(13) = "高雄" : arrChannels(13) = "網路商城" : arrAmounts(13) = 125000
arrZones(14) = "南部" : arrCities(14) = "高雄" : arrChannels(14) = "代理商"   : arrAmounts(14) = 92000
arrZones(15) = "南部" : arrCities(15) = "台南" : arrChannels(15) = "直營門市" : arrAmounts(15) = 118000
arrZones(16) = "南部" : arrCities(16) = "台南" : arrChannels(16) = "網路商城" : arrAmounts(16) = 86000
arrZones(17) = "南部" : arrCities(17) = "屏東" : arrChannels(17) = "直營門市" : arrAmounts(17) = 67000
arrZones(18) = "南部" : arrCities(18) = "屏東" : arrChannels(18) = "代理商"   : arrAmounts(18) = 45000
arrZones(19) = "東部" : arrCities(19) = "花蓮" : arrChannels(19) = "直營門市" : arrAmounts(19) = 54000
arrZones(20) = "東部" : arrCities(20) = "花蓮" : arrChannels(20) = "代理商"   : arrAmounts(20) = 38000
arrZones(21) = "東部" : arrCities(21) = "台東" : arrChannels(21) = "直營門市" : arrAmounts(21) = 42000
arrZones(22) = "東部" : arrCities(22) = "台東" : arrChannels(22) = "代理商"   : arrAmounts(22) = 31000
arrZones(23) = "北部" : arrCities(23) = "基隆" : arrChannels(23) = "直營門市" : arrAmounts(23) = 61000

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
objDataSheet.Cells(1, 1).Value = "大區"
objDataSheet.Cells(1, 2).Value = "城市"
objDataSheet.Cells(1, 3).Value = "通路"
objDataSheet.Cells(1, 4).Value = "銷售額"

With objDataSheet.Range("A1:D1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 23
    objDataSheet.Cells(i + 2, 1).Value = arrZones(i)
    objDataSheet.Cells(i + 2, 2).Value = arrCities(i)
    objDataSheet.Cells(i + 2, 3).Value = arrChannels(i)
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

' ── 設定三層巢狀列欄位（大區 > 城市 > 通路）─────────────────
Set objField = objPivot.PivotFields("大區")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("城市")
objField.Orientation = xlRowField
objField.Position    = 2

Set objField = objPivot.PivotFields("通路")
objField.Orientation = xlRowField
objField.Position    = 3

' ── 設定值欄位 ──────────────────────────────────────────────
Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 銷售額"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "三層巢狀列欄位樞紐分析表：大區 ＞ 城市 ＞ 通路"
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
