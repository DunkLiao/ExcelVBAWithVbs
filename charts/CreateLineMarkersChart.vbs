' ============================================================
' CreateLineMarkersChart.vbs
' 說明：使用 VBScript 自動建立 Excel 含資料點折線圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（雙城市全年氣溫對比）
'   3. 插入含資料標記點的折線圖
'   4. 設定圖表標題、座標軸標籤、圖例等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreateLineMarkersChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  = "2025 年台北 vs 高雄月均溫對比"
Const X_AXIS_TITLE = "月份"
Const Y_AXIS_TITLE = "溫度（°C）"
Const SHEET_NAME   = "氣溫對比"
Const OUTPUT_FILE  = "LineMarkersChartExample.xlsx"

' xlLineMarkers = 65（含資料點折線圖）
Const xlLineMarkers = 65
Const xlCategory    = 1
Const xlValue       = 2

' ── 範例資料 ────────────────────────────────────────────────
Dim arrMonths(11)
arrMonths(0)  = "1月"  : arrMonths(1)  = "2月"  : arrMonths(2)  = "3月"
arrMonths(3)  = "4月"  : arrMonths(4)  = "5月"  : arrMonths(5)  = "6月"
arrMonths(6)  = "7月"  : arrMonths(7)  = "8月"  : arrMonths(8)  = "9月"
arrMonths(9)  = "10月" : arrMonths(10) = "11月" : arrMonths(11) = "12月"

' 台北月均溫
Dim arrTaipei(11)
arrTaipei(0)  = 16 : arrTaipei(1)  = 17 : arrTaipei(2)  = 20
arrTaipei(3)  = 24 : arrTaipei(4)  = 28 : arrTaipei(5)  = 31
arrTaipei(6)  = 34 : arrTaipei(7)  = 33 : arrTaipei(8)  = 30
arrTaipei(9)  = 26 : arrTaipei(10) = 21 : arrTaipei(11) = 17

' 高雄月均溫
Dim arrKaohsiung(11)
arrKaohsiung(0)  = 20 : arrKaohsiung(1)  = 21 : arrKaohsiung(2)  = 24
arrKaohsiung(3)  = 27 : arrKaohsiung(4)  = 30 : arrKaohsiung(5)  = 32
arrKaohsiung(6)  = 33 : arrKaohsiung(7)  = 33 : arrKaohsiung(8)  = 31
arrKaohsiung(9)  = 28 : arrKaohsiung(10) = 24 : arrKaohsiung(11) = 21

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
objSheet.Cells(1, 2).Value = "台北（°C）"
objSheet.Cells(1, 3).Value = "高雄（°C）"

With objSheet.Range("A1:C1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 11
    objSheet.Cells(i + 2, 1).Value = arrMonths(i)
    objSheet.Cells(i + 2, 2).Value = arrTaipei(i)
    objSheet.Cells(i + 2, 3).Value = arrKaohsiung(i)
Next

objSheet.Columns("A:C").AutoFit()

' ── 插入含資料點折線圖 ───────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(230, 20, 480, 300)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlLineMarkers
objChart.SetSourceData objSheet.Range("A1:C13")

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

objChart.HasLegend = True

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
