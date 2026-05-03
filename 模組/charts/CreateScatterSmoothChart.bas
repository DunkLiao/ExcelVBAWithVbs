Attribute VB_Name = "CreateScatterSmoothChart"
' ============================================================
' CreateScatterSmoothChart.bas
' 說明：使用 Excel VBA 自動建立帶平滑線散佈圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（產品使用時數與效能衰減）
'   3. 插入帶平滑曲線的散佈圖
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateScatterSmoothChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  As String = "產品使用時數與效能衰減曲線"
Const X_AXIS_TITLE As String = "累計使用時數（小時）"
Const Y_AXIS_TITLE As String = "效能保留率（%）"
Const SHEET_NAME   As String = "效能測試"
Const OUTPUT_FILE  As String = "ScatterSmoothChartExample.xlsx"

Sub CreateScatterSmoothChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    ' 使用時數（小時）
    Dim arrHours(9) As Long
    arrHours(0) = 0    : arrHours(1) = 200  : arrHours(2) = 500
    arrHours(3) = 800  : arrHours(4) = 1000 : arrHours(5) = 1500
    arrHours(6) = 2000 : arrHours(7) = 2500 : arrHours(8) = 3000
    arrHours(9) = 3500

    ' 對應效能保留率（%）
    Dim arrPerf(9) As Long
    arrPerf(0) = 100 : arrPerf(1) = 98 : arrPerf(2) = 95
    arrPerf(3) = 91  : arrPerf(4) = 87 : arrPerf(5) = 80
    arrPerf(6) = 72  : arrPerf(7) = 63 : arrPerf(8) = 55
    arrPerf(9) = 48

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook As Workbook
    Dim objSheet    As Worksheet
    Dim objChartObj As ChartObject
    Dim objChart    As Chart
    Dim savePath    As String
    Dim i           As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook = Workbooks.Add()
    Set objSheet    = objWorkbook.Sheets(1)
    objSheet.Name   = SHEET_NAME

    ' ── 寫入標題列 ──────────────────────────────────────────────
    objSheet.Cells(1, 1).Value = "使用時數（hr）"
    objSheet.Cells(1, 2).Value = "效能保留率（%）"

    With objSheet.Range("A1:B1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 9
        objSheet.Cells(i + 2, 1).Value = arrHours(i)
        objSheet.Cells(i + 2, 2).Value = arrPerf(i)
    Next i

    objSheet.Columns("A:B").AutoFit

    ' ── 插入帶平滑線散佈圖 ───────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(200, 20, 480, 300)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlXYScatterSmooth
    objChart.SetSourceData objSheet.Range("A1:B11")

    ' ── 圖表格式設定 ────────────────────────────────────────────
    objChart.HasTitle        = True
    objChart.ChartTitle.Text = CHART_TITLE
    objChart.ChartTitle.Font.Size = 14
    objChart.ChartTitle.Font.Bold = True

    With objChart.Axes(xlCategory)
        .HasTitle       = True
        .AxisTitle.Text = X_AXIS_TITLE
        .AxisTitle.Font.Size = 10
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
    End With

    With objChart.Axes(xlValue)
        .HasTitle       = True
        .AxisTitle.Text = Y_AXIS_TITLE
        .AxisTitle.Font.Size = 10
        .MinimumScale = 0
        .MaximumScale = 100
    End With

    objChart.HasLegend = False

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objChart    = Nothing
    Set objChartObj = Nothing
    Set objSheet    = Nothing
    Set objWorkbook = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
