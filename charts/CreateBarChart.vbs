' ============================================================
' CreateBarChart.vbs
' 說明：使用 VBScript 自動建立 Excel 群組橫條圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（各部門人數）
'   3. 插入群組橫條圖
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreateBarChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  = "各部門員工人數"
Const X_AXIS_TITLE = "人數"
Const Y_AXIS_TITLE = "部門"
Const SHEET_NAME   = "部門人數"
Const OUTPUT_FILE  = "BarChartHorizExample.xlsx"

' xlBarClustered = 57（群組橫條圖）
Const xlBarClustered = 57
Const xlCategory     = 1
Const xlValue        = 2

' ── 範例資料 ────────────────────────────────────────────────
Dim arrDepts(6)
arrDepts(0) = "研發部" : arrDepts(1) = "業務部" : arrDepts(2) = "行銷部"
arrDepts(3) = "財務部" : arrDepts(4) = "人資部" : arrDepts(5) = "資訊部"
arrDepts(6) = "客服部"

Dim arrCount(6)
arrCount(0) = 45 : arrCount(1) = 32 : arrCount(2) = 18
arrCount(3) = 12 : arrCount(4) = 10 : arrCount(5) = 25
arrCount(6) = 20

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objSheet, objChartObj, objChart
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
objSheet.Cells(1, 1).Value = "部門"
objSheet.Cells(1, 2).Value = "人數"

With objSheet.Range("A1:B1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 6
    objSheet.Cells(i + 2, 1).Value = arrDepts(i)
    objSheet.Cells(i + 2, 2).Value = arrCount(i)
Next

objSheet.Columns("A:B").AutoFit()

' ── 插入群組橫條圖 ───────────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(200, 20, 480, 300)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlBarClustered
objChart.SetSourceData objSheet.Range("A1:B8")

' ── 圖表格式設定 ────────────────────────────────────────────
objChart.HasTitle        = True
objChart.ChartTitle.Text = CHART_TITLE
objChart.ChartTitle.Font.Size = 14
objChart.ChartTitle.Font.Bold = True

With objChart.Axes(xlCategory)
    .HasTitle       = True
    .AxisTitle.Text = Y_AXIS_TITLE
    .AxisTitle.Font.Size = 10
End With

With objChart.Axes(xlValue)
    .HasTitle       = True
    .AxisTitle.Text = X_AXIS_TITLE
    .AxisTitle.Font.Size = 10
    .MinimumScaleIsAuto = True
    .MaximumScaleIsAuto = True
End With

objChart.SeriesCollection(1).HasDataLabels = True
objChart.HasLegend = False

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objChart    = Nothing
Set objChartObj = Nothing
Set objSheet    = Nothing
Set objWorkbook = Nothing
Set objExcel    = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
