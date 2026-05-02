' ============================================================
' CreateStackedBarChart.vbs
' 說明：使用 VBScript 自動建立 Excel 堆疊橫條圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（各地區三項產品銷售堆疊）
'   3. 插入堆疊橫條圖
'   4. 設定圖表標題、座標軸標籤、圖例等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreateStackedBarChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  = "各地區產品銷售堆疊（萬元）"
Const X_AXIS_TITLE = "銷售額（萬元）"
Const Y_AXIS_TITLE = "地區"
Const SHEET_NAME   = "地區銷售"
Const OUTPUT_FILE  = "StackedBarChartExample.xlsx"

' xlBarStacked = 58（堆疊橫條圖）
Const xlBarStacked = 58
Const xlCategory   = 1
Const xlValue      = 2

' ── 範例資料 ────────────────────────────────────────────────
' 欄位：地區, 產品A, 產品B, 產品C
Dim arrRows(4, 3)
arrRows(0, 0) = "北部" : arrRows(0, 1) = 180 : arrRows(0, 2) = 120 : arrRows(0, 3) = 90
arrRows(1, 0) = "中部" : arrRows(1, 1) = 150 : arrRows(1, 2) = 100 : arrRows(1, 3) = 70
arrRows(2, 0) = "南部" : arrRows(2, 1) = 200 : arrRows(2, 2) = 140 : arrRows(2, 3) = 110
arrRows(3, 0) = "東部" : arrRows(3, 1) =  80 : arrRows(3, 2) =  60 : arrRows(3, 3) = 40
arrRows(4, 0) = "離島" : arrRows(4, 1) =  50 : arrRows(4, 2) =  30 : arrRows(4, 3) = 25

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objSheet, objChartObj, objChart
Dim savePath, objShell, r, c

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
objSheet.Cells(1, 1).Value = "地區"
objSheet.Cells(1, 2).Value = "產品A"
objSheet.Cells(1, 3).Value = "產品B"
objSheet.Cells(1, 4).Value = "產品C"

With objSheet.Range("A1:D1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For r = 0 To 4
    For c = 0 To 3
        objSheet.Cells(r + 2, c + 1).Value = arrRows(r, c)
    Next
Next

objSheet.Columns("A:D").AutoFit()

' ── 插入堆疊橫條圖 ───────────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(260, 20, 480, 300)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlBarStacked
objChart.SetSourceData objSheet.Range("A1:D6")

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
