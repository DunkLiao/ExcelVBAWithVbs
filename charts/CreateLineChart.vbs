' ============================================================
' CreateLineChart.vbs
' 說明：使用 VBScript 自動建立 Excel 折線圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（全年平均氣溫）
'   3. 插入折線圖
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreateLineChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  = "2025 年各月平均氣溫"
Const X_AXIS_TITLE = "月份"
Const Y_AXIS_TITLE = "溫度（°C）"
Const SHEET_NAME   = "氣溫資料"
Const OUTPUT_FILE  = "LineChartExample.xlsx"

' xlLine = 4（折線圖）
Const xlLine     = 4
Const xlCategory = 1
Const xlValue    = 2

' ── 範例資料 ────────────────────────────────────────────────
Dim arrMonths(11)
arrMonths(0)  = "1月"  : arrMonths(1)  = "2月"  : arrMonths(2)  = "3月"
arrMonths(3)  = "4月"  : arrMonths(4)  = "5月"  : arrMonths(5)  = "6月"
arrMonths(6)  = "7月"  : arrMonths(7)  = "8月"  : arrMonths(8)  = "9月"
arrMonths(9)  = "10月" : arrMonths(10) = "11月" : arrMonths(11) = "12月"

Dim arrTemp(11)
arrTemp(0)  = 15 : arrTemp(1)  = 16 : arrTemp(2)  = 19
arrTemp(3)  = 23 : arrTemp(4)  = 27 : arrTemp(5)  = 31
arrTemp(6)  = 34 : arrTemp(7)  = 33 : arrTemp(8)  = 29
arrTemp(9)  = 25 : arrTemp(10) = 20 : arrTemp(11) = 16

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
objSheet.Cells(1, 1).Value = "月份"
objSheet.Cells(1, 2).Value = "平均氣溫（°C）"

With objSheet.Range("A1:B1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 11
    objSheet.Cells(i + 2, 1).Value = arrMonths(i)
    objSheet.Cells(i + 2, 2).Value = arrTemp(i)
Next

objSheet.Columns("A:B").AutoFit()

' ── 插入折線圖 ───────────────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(200, 20, 480, 300)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlLine
objChart.SetSourceData objSheet.Range("A1:B13")

' ── 圖表格式設定 ────────────────────────────────────────────
objChart.HasTitle        = True
objChart.ChartTitle.Text = CHART_TITLE
objChart.ChartTitle.Font.Size = 14
objChart.ChartTitle.Font.Bold = True

With objChart.Axes(xlCategory)
    .HasTitle       = True
    .AxisTitle.Text = X_AXIS_TITLE
    .AxisTitle.Font.Size = 10
End With

With objChart.Axes(xlValue)
    .HasTitle       = True
    .AxisTitle.Text = Y_AXIS_TITLE
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
