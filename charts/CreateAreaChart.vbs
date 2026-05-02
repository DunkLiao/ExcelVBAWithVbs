' ============================================================
' CreateAreaChart.vbs
' 說明：使用 VBScript 自動建立 Excel 區域圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（網站月流量）
'   3. 插入區域圖
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreateAreaChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  = "2025 年網站月流量趨勢"
Const X_AXIS_TITLE = "月份"
Const Y_AXIS_TITLE = "瀏覽次數（千次）"
Const SHEET_NAME   = "流量資料"
Const OUTPUT_FILE  = "AreaChartExample.xlsx"

' xlArea = 1（區域圖）
Const xlArea     = 1
Const xlCategory = 1
Const xlValue    = 2

' ── 範例資料 ────────────────────────────────────────────────
Dim arrMonths(11)
arrMonths(0)  = "1月"  : arrMonths(1)  = "2月"  : arrMonths(2)  = "3月"
arrMonths(3)  = "4月"  : arrMonths(4)  = "5月"  : arrMonths(5)  = "6月"
arrMonths(6)  = "7月"  : arrMonths(7)  = "8月"  : arrMonths(8)  = "9月"
arrMonths(9)  = "10月" : arrMonths(10) = "11月" : arrMonths(11) = "12月"

Dim arrViews(11)
arrViews(0)  = 320 : arrViews(1)  = 280 : arrViews(2)  = 410
arrViews(3)  = 500 : arrViews(4)  = 620 : arrViews(5)  = 590
arrViews(6)  = 680 : arrViews(7)  = 710 : arrViews(8)  = 650
arrViews(9)  = 730 : arrViews(10) = 820 : arrViews(11) = 950

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
objSheet.Cells(1, 2).Value = "瀏覽次數（千次）"

With objSheet.Range("A1:B1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 11
    objSheet.Cells(i + 2, 1).Value = arrMonths(i)
    objSheet.Cells(i + 2, 2).Value = arrViews(i)
Next

objSheet.Columns("A:B").AutoFit()

' ── 插入區域圖 ───────────────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(200, 20, 480, 300)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlArea
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
