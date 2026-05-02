' ============================================================
' CreateColumnStacked100Chart.vbs
' 說明：使用 VBScript 自動建立 Excel 百分比堆疊直條圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（各季度市場佔有率變化）
'   3. 插入百分比堆疊直條圖（每欄合計 = 100%）
'   4. 設定圖表標題、座標軸標籤、圖例等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreateColumnStacked100Chart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  = "2025 年各季品牌市佔率變化（%）"
Const X_AXIS_TITLE = "季度"
Const Y_AXIS_TITLE = "市佔率（%）"
Const SHEET_NAME   = "市佔率"
Const OUTPUT_FILE  = "ColumnStacked100ChartExample.xlsx"

' xlColumnStacked100 = 53（百分比堆疊直條圖）
Const xlColumnStacked100 = 53
Const xlCategory         = 1
Const xlValue            = 2

' ── 範例資料 ────────────────────────────────────────────────
' 欄位：季度, 品牌A, 品牌B, 品牌C, 其他
Dim arrRows(3, 4)
arrRows(0, 0) = "Q1" : arrRows(0, 1) = 38 : arrRows(0, 2) = 28 : arrRows(0, 3) = 20 : arrRows(0, 4) = 14
arrRows(1, 0) = "Q2" : arrRows(1, 1) = 40 : arrRows(1, 2) = 27 : arrRows(1, 3) = 20 : arrRows(1, 4) = 13
arrRows(2, 0) = "Q3" : arrRows(2, 1) = 42 : arrRows(2, 2) = 25 : arrRows(2, 3) = 21 : arrRows(2, 4) = 12
arrRows(3, 0) = "Q4" : arrRows(3, 1) = 45 : arrRows(3, 2) = 24 : arrRows(3, 3) = 20 : arrRows(3, 4) = 11

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
objSheet.Cells(1, 1).Value = "季度"
objSheet.Cells(1, 2).Value = "品牌A"
objSheet.Cells(1, 3).Value = "品牌B"
objSheet.Cells(1, 4).Value = "品牌C"
objSheet.Cells(1, 5).Value = "其他"

With objSheet.Range("A1:E1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For r = 0 To 3
    For c = 0 To 4
        objSheet.Cells(r + 2, c + 1).Value = arrRows(r, c)
    Next
Next

objSheet.Columns("A:E").AutoFit()

' ── 插入百分比堆疊直條圖 ─────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(280, 20, 480, 300)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlColumnStacked100
objChart.SetSourceData objSheet.Range("A1:E5")

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
    .MinimumScale = 0
    .MaximumScale = 100
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
