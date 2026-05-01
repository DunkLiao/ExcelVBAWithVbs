' ============================================================
' CreateBarChart.vbs
' 說明：使用 VBScript 自動建立 Excel 長條圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（月份銷售額）
'   3. 根據資料插入群組直條圖（長條圖）
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript CreateBarChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE    = "2025 年各月銷售額"
Const X_AXIS_TITLE   = "月份"
Const Y_AXIS_TITLE   = "銷售額（萬元）"
Const SHEET_NAME     = "銷售資料"
Const OUTPUT_FILE    = "BarChartExample.xlsx"   ' 輸出至桌面

' xlClusteredColumn = 51（群組直條圖）
Const xlClusteredColumn = 51
' xlCategory = 1, xlValue = 2
Const xlCategory = 1
Const xlValue    = 2

' ── 範例資料 ────────────────────────────────────────────────
' 月份標題
Dim arrMonths(11)
arrMonths(0)  = "1月"  : arrMonths(1)  = "2月"  : arrMonths(2)  = "3月"
arrMonths(3)  = "4月"  : arrMonths(4)  = "5月"  : arrMonths(5)  = "6月"
arrMonths(6)  = "7月"  : arrMonths(7)  = "8月"  : arrMonths(8)  = "9月"
arrMonths(9)  = "10月" : arrMonths(10) = "11月" : arrMonths(11) = "12月"

' 各月銷售額（萬元）
Dim arrSales(11)
arrSales(0)  = 120 : arrSales(1)  = 95  : arrSales(2)  = 150
arrSales(3)  = 180 : arrSales(4)  = 210 : arrSales(5)  = 230
arrSales(6)  = 200 : arrSales(7)  = 175 : arrSales(8)  = 195
arrSales(9)  = 250 : arrSales(10) = 300 : arrSales(11) = 400

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objSheet, objChartObj, objChart
Dim i, lastRow, dataRange, savePath

' 取得桌面路徑
Dim objShell
Set objShell = CreateObject("WScript.Shell")
savePath = objShell.SpecialFolders("Desktop") & "\" & OUTPUT_FILE
Set objShell = Nothing

' 建立 Excel 物件
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False        ' 背景執行；若想看過程可改為 True
objExcel.DisplayAlerts = False

' 新增活頁簿
Set objWorkbook = objExcel.Workbooks.Add()
Set objSheet    = objWorkbook.Sheets(1)
objSheet.Name   = SHEET_NAME

' ── 寫入標題列 ──────────────────────────────────────────────
objSheet.Cells(1, 1).Value = "月份"
objSheet.Cells(1, 2).Value = "銷售額（萬元）"

' 設定標題列格式：粗體 + 置中
With objSheet.Range("A1:B1")
    .Font.Bold    = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 11
    objSheet.Cells(i + 2, 1).Value = arrMonths(i)   ' A 欄：月份
    objSheet.Cells(i + 2, 2).Value = arrSales(i)    ' B 欄：銷售額
Next

lastRow   = 13   ' 資料最後一列（標題列 + 12 個月）
dataRange = "A1:B" & lastRow

' 自動調整欄寬
objSheet.Columns("A:B").AutoFit()

' ── 插入長條圖（群組直條圖）────────────────────────────────
' 在工作表上建立內嵌圖表物件，位置：左上角 D2，大小 480×300 點
Set objChartObj = objSheet.ChartObjects.Add(200, 20, 480, 300)
Set objChart    = objChartObj.Chart

' 設定圖表類型：群組直條圖
objChart.ChartType = xlClusteredColumn

' 指定資料來源（含標題列，Excel 自動辨識類別與數值）
objChart.SetSourceData objSheet.Range(dataRange)

' ── 圖表格式設定 ────────────────────────────────────────────
' 圖表標題
objChart.HasTitle           = True
objChart.ChartTitle.Text    = CHART_TITLE
objChart.ChartTitle.Font.Size = 14
objChart.ChartTitle.Font.Bold = True

' X 軸（類別軸）標題
With objChart.Axes(xlCategory)
    .HasTitle          = True
    .AxisTitle.Text    = X_AXIS_TITLE
    .AxisTitle.Font.Size = 10
End With

' Y 軸（數值軸）標題
With objChart.Axes(xlValue)
    .HasTitle          = True
    .AxisTitle.Text    = Y_AXIS_TITLE
    .AxisTitle.Font.Size = 10
    .MinimumScaleIsAuto = True
    .MaximumScaleIsAuto = True
End With

' 顯示資料標籤
objChart.SeriesCollection(1).HasDataLabels = True

' 移除圖例（只有一個數列時圖例意義不大）
objChart.HasLegend = False

' ── 儲存並關閉 ──────────────────────────────────────────────
' 51 = xlOpenXMLWorkbook (.xlsx)
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

' 釋放物件
Set objChart    = Nothing
Set objChartObj = Nothing
Set objSheet    = Nothing
Set objWorkbook = Nothing
Set objExcel    = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
