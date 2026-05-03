Attribute VB_Name = "CreateStackedColumnChart"
' ============================================================
' CreateStackedColumnChart.bas
' 說明：使用 Excel VBA 自動建立堆疊直條圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（季度收支堆疊）
'   3. 插入堆疊直條圖
'   4. 設定圖表標題、座標軸標籤、圖例等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateStackedColumnChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  As String = "2025 年季度收支分析"
Const X_AXIS_TITLE As String = "季度"
Const Y_AXIS_TITLE As String = "金額（萬元）"
Const SHEET_NAME   As String = "收支資料"
Const OUTPUT_FILE  As String = "StackedColumnChartExample.xlsx"

Sub CreateStackedColumnChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    ' 欄位：季度, 薪資, 租金, 行銷, 其他
    Dim arrRows(3, 4) As Variant
    arrRows(0, 0) = "Q1" : arrRows(0, 1) = 200 : arrRows(0, 2) = 50 : arrRows(0, 3) = 80  : arrRows(0, 4) = 30
    arrRows(1, 0) = "Q2" : arrRows(1, 1) = 210 : arrRows(1, 2) = 50 : arrRows(1, 3) = 95  : arrRows(1, 4) = 35
    arrRows(2, 0) = "Q3" : arrRows(2, 1) = 220 : arrRows(2, 2) = 50 : arrRows(2, 3) = 110 : arrRows(2, 4) = 40
    arrRows(3, 0) = "Q4" : arrRows(3, 1) = 250 : arrRows(3, 2) = 50 : arrRows(3, 3) = 150 : arrRows(3, 4) = 60

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

    ' ── 寫入標題列 ──────────────────────────────────────────────
    objSheet.Cells(1, 1).Value = "季度"
    objSheet.Cells(1, 2).Value = "薪資成本"
    objSheet.Cells(1, 3).Value = "租金"
    objSheet.Cells(1, 4).Value = "行銷費用"
    objSheet.Cells(1, 5).Value = "其他費用"

    With objSheet.Range("A1:E1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For r = 0 To 3
        For c = 0 To 4
            objSheet.Cells(r + 2, c + 1).Value = arrRows(r, c)
        Next c
    Next r

    objSheet.Columns("A:E").AutoFit

    ' ── 插入堆疊直條圖 ───────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(280, 20, 480, 300)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlColumnStacked
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
