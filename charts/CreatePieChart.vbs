' ============================================================
' CreatePieChart.vbs
' 說明：使用 VBScript 自動建立 Excel 圓餅圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（市場佔有率）
'   3. 插入圓餅圖
'   4. 設定圖表標題、百分比資料標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreatePieChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE = "2025 年智慧型手機市場佔有率"
Const SHEET_NAME  = "市場佔有率"
Const OUTPUT_FILE = "PieChartExample.xlsx"

' xlPie = 5（圓餅圖）
Const xlPie = 5

' ── 範例資料 ────────────────────────────────────────────────
Dim arrBrands(4)
arrBrands(0) = "品牌A" : arrBrands(1) = "品牌B" : arrBrands(2) = "品牌C"
arrBrands(3) = "品牌D" : arrBrands(4) = "其他"

Dim arrShare(4)
arrShare(0) = 32 : arrShare(1) = 25 : arrShare(2) = 18
arrShare(3) = 12 : arrShare(4) = 13

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objSheet, objChartObj, objChart
Dim objLabel, savePath, objShell, i

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
objSheet.Cells(1, 1).Value = "品牌"
objSheet.Cells(1, 2).Value = "市佔率（%）"

With objSheet.Range("A1:B1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 4
    objSheet.Cells(i + 2, 1).Value = arrBrands(i)
    objSheet.Cells(i + 2, 2).Value = arrShare(i)
Next

objSheet.Columns("A:B").AutoFit()

' ── 插入圓餅圖 ───────────────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(200, 20, 400, 300)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlPie
objChart.SetSourceData objSheet.Range("A1:B6")

' ── 圖表格式設定 ────────────────────────────────────────────
objChart.HasTitle        = True
objChart.ChartTitle.Text = CHART_TITLE
objChart.ChartTitle.Font.Size = 14
objChart.ChartTitle.Font.Bold = True

' 顯示百分比資料標籤 + 類別名稱
With objChart.SeriesCollection(1)
    .HasDataLabels = True
    With .DataLabels
        .ShowPercentage  = True
        .ShowCategoryName = True
        .ShowValue        = False
    End With
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
