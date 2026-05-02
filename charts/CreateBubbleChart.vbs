' ============================================================
' CreateBubbleChart.vbs
' 說明：使用 VBScript 自動建立 Excel 泡泡圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（城市 GDP / 人口 / 幸福指數）
'   3. 插入泡泡圖（X=GDP, Y=人口, 泡泡大小=幸福指數）
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreateBubbleChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  = "城市 GDP、人口與幸福指數"
Const X_AXIS_TITLE = "GDP（千億元）"
Const Y_AXIS_TITLE = "人口（萬人）"
Const SHEET_NAME   = "城市資料"
Const OUTPUT_FILE  = "BubbleChartExample.xlsx"

' xlBubble = 15（泡泡圖）
Const xlBubble   = 15
Const xlCategory = 1
Const xlValue    = 2

' ── 範例資料 ────────────────────────────────────────────────
' 欄位：城市, GDP（千億元）, 人口（萬人）, 幸福指數（泡泡大小）
Dim arrCity(5)
arrCity(0) = "台北" : arrCity(1) = "新北" : arrCity(2) = "桃園"
arrCity(3) = "台中" : arrCity(4) = "台南" : arrCity(5) = "高雄"

Dim arrGDP(5)
arrGDP(0) = 38 : arrGDP(1) = 22 : arrGDP(2) = 18
arrGDP(3) = 20 : arrGDP(4) = 12 : arrGDP(5) = 16

Dim arrPop(5)
arrPop(0) = 267 : arrPop(1) = 403 : arrPop(2) = 229
arrPop(3) = 281 : arrPop(4) = 188 : arrPop(5) = 276

Dim arrHappy(5)
arrHappy(0) = 75 : arrHappy(1) = 70 : arrHappy(2) = 72
arrHappy(3) = 78 : arrHappy(4) = 80 : arrHappy(5) = 74

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objSheet, objChartObj, objChart, objSeries
Dim savePath, objShell, i

Set objShell = CreateObject("WScript.Shell")
savePath = objShell.SpecialFolders("Desktop") & "\" & OUTPUT_FILE
Set objShell = Nothing

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible       = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Add()
Set objSheet    = objWorkbook.Sheets(1)
objSheet.Name   = SHEET_NAME

' ── 寫入標題列 ──────────────────────────────────────────────
objSheet.Cells(1, 1).Value = "城市"
objSheet.Cells(1, 2).Value = "GDP（千億元）"
objSheet.Cells(1, 3).Value = "人口（萬人）"
objSheet.Cells(1, 4).Value = "幸福指數"

With objSheet.Range("A1:D1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 5
    objSheet.Cells(i + 2, 1).Value = arrCity(i)
    objSheet.Cells(i + 2, 2).Value = arrGDP(i)
    objSheet.Cells(i + 2, 3).Value = arrPop(i)
    objSheet.Cells(i + 2, 4).Value = arrHappy(i)
Next

objSheet.Columns("A:D").AutoFit()

' ── 插入泡泡圖 ───────────────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(270, 20, 480, 300)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlBubble

' 泡泡圖需手動指定 X、Y 與泡泡大小數列
objChart.SeriesCollection.NewSeries()
Set objSeries = objChart.SeriesCollection(1)
objSeries.Name          = "城市"
objSeries.XValues       = objSheet.Range("B2:B7")
objSeries.Values        = objSheet.Range("C2:C7")
objSeries.BubbleSizes   = objSheet.Range("D2:D7")

' ── 圖表格式設定 ────────────────────────────────────────────
objChart.HasTitle        = True
objChart.ChartTitle.Text = CHART_TITLE
objChart.ChartTitle.Font.Size = 14
objChart.ChartTitle.Font.Bold = True

With objChart.Axes(xlCategory)
    .HasTitle       = True
    .AxisTitle.Text = X_AXIS_TITLE
    .AxisTitle.Font.Size = 10
    .MinimumScaleIsAuto = True
    .MaximumScaleIsAuto = True
End With

With objChart.Axes(xlValue)
    .HasTitle       = True
    .AxisTitle.Text = Y_AXIS_TITLE
    .AxisTitle.Font.Size = 10
    .MinimumScaleIsAuto = True
    .MaximumScaleIsAuto = True
End With

objChart.HasLegend = False

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objSeries   = Nothing
Set objChart    = Nothing
Set objChartObj = Nothing
Set objSheet    = Nothing
Set objWorkbook = Nothing
Set objExcel    = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
