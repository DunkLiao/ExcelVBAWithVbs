Attribute VB_Name = "CreateLineChart"
' ============================================================
' CreateLineChart.bas
' 說明：使用 Excel VBA 自動建立折線圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（全年平均氣溫）
'   3. 插入折線圖
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateLineChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  As String = "2025 年各月平均氣溫"
Const X_AXIS_TITLE As String = "月份"
Const Y_AXIS_TITLE As String = "溫度（°C）"
Const SHEET_NAME   As String = "氣溫資料"
Const OUTPUT_FILE  As String = "LineChartExample.xlsx"

Sub CreateLineChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    Dim arrMonths(11) As String
    arrMonths(0)  = "1月"  : arrMonths(1)  = "2月"  : arrMonths(2)  = "3月"
    arrMonths(3)  = "4月"  : arrMonths(4)  = "5月"  : arrMonths(5)  = "6月"
    arrMonths(6)  = "7月"  : arrMonths(7)  = "8月"  : arrMonths(8)  = "9月"
    arrMonths(9)  = "10月" : arrMonths(10) = "11月" : arrMonths(11) = "12月"

    Dim arrTemp(11) As Long
    arrTemp(0)  = 15 : arrTemp(1)  = 16 : arrTemp(2)  = 19
    arrTemp(3)  = 23 : arrTemp(4)  = 27 : arrTemp(5)  = 31
    arrTemp(6)  = 34 : arrTemp(7)  = 33 : arrTemp(8)  = 29
    arrTemp(9)  = 25 : arrTemp(10) = 20 : arrTemp(11) = 16

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
    objSheet.Cells(1, 1).Value = "月份"
    objSheet.Cells(1, 2).Value = "平均氣溫（°C）"

    With objSheet.Range("A1:B1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 11
        objSheet.Cells(i + 2, 1).Value = arrMonths(i)
        objSheet.Cells(i + 2, 2).Value = arrTemp(i)
    Next i

    objSheet.Columns("A:B").AutoFit

    ' ── 插入折線圖 ───────────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(200, 20, 480, 300)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlLine
    objChart.SetSourceData objSheet.Range("A1:B13")

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

    objChart.SeriesCollection(1).HasDataLabels = True
    objChart.HasLegend = False

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objChart    = Nothing
    Set objChartObj = Nothing
    Set objSheet    = Nothing
    Set objWorkbook = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
