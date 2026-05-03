Attribute VB_Name = "CreateSurfaceChart"
' ============================================================
' CreateSurfaceChart.bas
' 說明：使用 Excel VBA 自動建立曲面圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（溫度與壓力對應數值矩陣）
'   3. 插入曲面圖（3D 等高線圖）
'   4. 設定圖表標題等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateSurfaceChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE As String = "溫度與壓力條件下的反應速率"
Const SHEET_NAME  As String = "反應速率矩陣"
Const OUTPUT_FILE As String = "SurfaceChartExample.xlsx"

Sub CreateSurfaceChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    ' 列標籤：溫度（°C）100, 150, 200, 250, 300
    ' 欄標籤：壓力（atm）1, 2, 3, 4, 5
    ' 矩陣中值：反應速率（相對單位）
    Dim arrTemp(4) As String
    arrTemp(0) = "100°C" : arrTemp(1) = "150°C" : arrTemp(2) = "200°C"
    arrTemp(3) = "250°C" : arrTemp(4) = "300°C"

    ' 反應速率矩陣（溫度列 × 壓力欄）
    Dim arrRate(4, 4) As Long
    arrRate(0, 0) = 10  : arrRate(0, 1) = 15  : arrRate(0, 2) = 19  : arrRate(0, 3) = 22  : arrRate(0, 4) = 24
    arrRate(1, 0) = 20  : arrRate(1, 1) = 28  : arrRate(1, 2) = 35  : arrRate(1, 3) = 40  : arrRate(1, 4) = 44
    arrRate(2, 0) = 35  : arrRate(2, 1) = 48  : arrRate(2, 2) = 58  : arrRate(2, 3) = 65  : arrRate(2, 4) = 70
    arrRate(3, 0) = 55  : arrRate(3, 1) = 72  : arrRate(3, 2) = 85  : arrRate(3, 3) = 94  : arrRate(3, 4) = 100
    arrRate(4, 0) = 80  : arrRate(4, 1) = 100 : arrRate(4, 2) = 115 : arrRate(4, 3) = 126 : arrRate(4, 4) = 134

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook As Workbook
    Dim objSheet    As Worksheet
    Dim objChartObj As ChartObject
    Dim objChart    As Chart
    Dim savePath    As String
    Dim r           As Integer
    Dim c           As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook = Workbooks.Add()
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
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入列標題（溫度 °C）與矩陣資料 ─────────────────────────
    For r = 0 To 4
        objSheet.Cells(r + 2, 1).Value = arrTemp(r)
        For c = 0 To 4
            objSheet.Cells(r + 2, c + 2).Value = arrRate(r, c)
        Next c
    Next r

    objSheet.Columns("A:F").AutoFit

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

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objChart    = Nothing
    Set objChartObj = Nothing
    Set objSheet    = Nothing
    Set objWorkbook = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
