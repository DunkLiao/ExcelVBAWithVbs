' ============================================================
' CreatePie3DChart.vbs
' 說明：使用 VBScript 自動建立 Excel 3D 圓餅圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（公司資源分配比例）
'   3. 插入 3D 圓餅圖
'   4. 設定圖表標題、百分比資料標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreatePie3DChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE = "2025 年公司資源分配（3D）"
Const SHEET_NAME  = "資源分配"
Const OUTPUT_FILE = "Pie3DChartExample.xlsx"

' xl3DPie = -4102（3D 圓餅圖）
Const xl3DPie = -4102

' ── 範例資料 ────────────────────────────────────────────────
Dim arrItems(4)
arrItems(0) = "技術研發" : arrItems(1) = "市場推廣" : arrItems(2) = "人才培育"
arrItems(3) = "基礎設施" : arrItems(4) = "客戶服務"

Dim arrPercent(4)
arrPercent(0) = 35 : arrPercent(1) = 25 : arrPercent(2) = 20
arrPercent(3) = 12 : arrPercent(4) =  8

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
objSheet.Cells(1, 1).Value = "資源項目"
objSheet.Cells(1, 2).Value = "佔比（%）"

With objSheet.Range("A1:B1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 4
    objSheet.Cells(i + 2, 1).Value = arrItems(i)
    objSheet.Cells(i + 2, 2).Value = arrPercent(i)
Next

objSheet.Columns("A:B").AutoFit()

' ── 插入 3D 圓餅圖 ───────────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(200, 20, 420, 310)
Set objChart    = objChartObj.Chart

objChart.ChartType = xl3DPie
objChart.SetSourceData objSheet.Range("A1:B6")

' ── 圖表格式設定 ────────────────────────────────────────────
objChart.HasTitle        = True
objChart.ChartTitle.Text = CHART_TITLE
objChart.ChartTitle.Font.Size = 14
objChart.ChartTitle.Font.Bold = True

' 顯示百分比 + 類別名稱
With objChart.SeriesCollection(1)
    .HasDataLabels = True
    With .DataLabels
        .ShowPercentage   = True
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
