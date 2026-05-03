Attribute VB_Name = "CreateColumnChart"
' ============================================================
' CreateColumnChart.bas
' 說明：使用 Excel VBA 自動建立群組直條圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（三個產品的季度業績）
'   3. 插入群組直條圖（多數列）
'   4. 設定圖表標題、座標軸標籤、圖例等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateColumnChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  As String = "2025 年各季產品業績比較"
Const X_AXIS_TITLE As String = "季度"
Const Y_AXIS_TITLE As String = "業績（萬元）"
Const SHEET_NAME   As String = "業績資料"
Const OUTPUT_FILE  As String = "ColumnChartExample.xlsx"

Sub CreateColumnChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    Dim arrHeaders(3) As String
    arrHeaders(0) = "季度" : arrHeaders(1) = "產品A" : arrHeaders(2) = "產品B" : arrHeaders(3) = "產品C"

    Dim arrRows(3, 3) As Variant
    arrRows(0, 0) = "Q1" : arrRows(0, 1) = 120 : arrRows(0, 2) = 95  : arrRows(0, 3) = 80
    arrRows(1, 0) = "Q2" : arrRows(1, 1) = 150 : arrRows(1, 2) = 130 : arrRows(1, 3) = 110
    arrRows(2, 0) = "Q3" : arrRows(2, 1) = 200 : arrRows(2, 2) = 160 : arrRows(2, 3) = 140
    arrRows(3, 0) = "Q4" : arrRows(3, 1) = 280 : arrRows(3, 2) = 220 : arrRows(3, 3) = 190

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook As Workbook
    Dim objSheet    As Worksheet
    Dim objChartObj As ChartObject
    Dim objChart    As Chart
    Dim savePath    As String
    Dim i           As Integer
    Dim r           As Integer
    Dim c           As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook = Workbooks.Add()
    Set objSheet    = objWorkbook.Sheets(1)
    objSheet.Name   = SHEET_NAME

    ' ── 寫入標題列 ──────────────────────────────────────────────
    For i = 0 To 3
        objSheet.Cells(1, i + 1).Value = arrHeaders(i)
    Next i

    With objSheet.Range("A1:D1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For r = 0 To 3
        For c = 0 To 3
            objSheet.Cells(r + 2, c + 1).Value = arrRows(r, c)
        Next c
    Next r

    objSheet.Columns("A:D").AutoFit

    ' ── 插入群組直條圖 ───────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(230, 20, 480, 300)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlClusteredColumn
    objChart.SetSourceData objSheet.Range("A1:D5")

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

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objChart    = Nothing
    Set objChartObj = Nothing
    Set objSheet    = Nothing
    Set objWorkbook = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
