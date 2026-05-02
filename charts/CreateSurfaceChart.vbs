' ============================================================
' CreateSurfaceChart.vbs
' 說明：使用 VBScript 自動建立 Excel 曲面圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（溫度與壓力對應數值矩陣）
'   3. 插入曲面圖（3D 等高線圖）
'   4. 設定圖表標題等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreateSurfaceChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE = "溫度與壓力條件下的反應速率"
Const SHEET_NAME  = "反應速率矩陣"
Const OUTPUT_FILE = "SurfaceChartExample.xlsx"

' xlSurface = 83（3D 曲面圖）
Const xlSurface = 83

' ── 範例資料 ────────────────────────────────────────────────
' 列標籤：溫度（°C）100, 150, 200, 250, 300
' 欄標籤：壓力（atm）1, 2, 3, 4, 5
' 矩陣中值：反應速率（相對單位）

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

' ── 寫入欄標題（壓力 atm）──────────────────────────────────
objSheet.Cells(1, 1).Value = "溫度\壓力"
objSheet.Cells(1, 2).Value = "1 atm"
objSheet.Cells(1, 3).Value = "2 atm"
objSheet.Cells(1, 4).Value = "3 atm"
objSheet.Cells(1, 5).Value = "4 atm"
objSheet.Cells(1, 6).Value = "5 atm"

With objSheet.Range("A1:F1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入列標題（溫度 °C）與矩陣資料 ─────────────────────────
Dim arrTemp(4)
arrTemp(0) = "100°C" : arrTemp(1) = "150°C" : arrTemp(2) = "200°C"
arrTemp(3) = "250°C" : arrTemp(4) = "300°C"

' 反應速率矩陣（溫度列 × 壓力欄）
Dim arrRate(4, 4)
arrRate(0, 0) = 10 : arrRate(0, 1) = 15 : arrRate(0, 2) = 19 : arrRate(0, 3) = 22 : arrRate(0, 4) = 24
arrRate(1, 0) = 20 : arrRate(1, 1) = 28 : arrRate(1, 2) = 35 : arrRate(1, 3) = 40 : arrRate(1, 4) = 44
arrRate(2, 0) = 35 : arrRate(2, 1) = 48 : arrRate(2, 2) = 58 : arrRate(2, 3) = 65 : arrRate(2, 4) = 70
arrRate(3, 0) = 55 : arrRate(3, 1) = 72 : arrRate(3, 2) = 85 : arrRate(3, 3) = 94 : arrRate(3, 4) = 100
arrRate(4, 0) = 80 : arrRate(4, 1) = 100 : arrRate(4, 2) = 115 : arrRate(4, 3) = 126 : arrRate(4, 4) = 134

For r = 0 To 4
    objSheet.Cells(r + 2, 1).Value = arrTemp(r)
    For c = 0 To 4
        objSheet.Cells(r + 2, c + 2).Value = arrRate(r, c)
    Next
Next

objSheet.Columns("A:F").AutoFit()

' ── 插入曲面圖 ───────────────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(280, 20, 480, 320)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlSurface
objChart.SetSourceData objSheet.Range("A1:F6")

' ── 圖表格式設定 ────────────────────────────────────────────
objChart.HasTitle        = True
objChart.ChartTitle.Text = CHART_TITLE
objChart.ChartTitle.Font.Size = 14
objChart.ChartTitle.Font.Bold = True

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
