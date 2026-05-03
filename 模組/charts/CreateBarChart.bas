Attribute VB_Name = "CreateBarChart"
' ============================================================
' CreateBarChart.bas
' 說明：使用 Excel VBA 自動建立群組橫條圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（各部門人數）
'   3. 插入群組橫條圖
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateBarChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  As String = "各部門員工人數"
Const X_AXIS_TITLE As String = "人數"
Const Y_AXIS_TITLE As String = "部門"
Const SHEET_NAME   As String = "部門人數"
Const OUTPUT_FILE  As String = "BarChartHorizExample.xlsx"

Sub CreateBarChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    Dim arrDepts(6) As String
    arrDepts(0) = "研發部" : arrDepts(1) = "業務部" : arrDepts(2) = "行銷部"
    arrDepts(3) = "財務部" : arrDepts(4) = "人資部" : arrDepts(5) = "資訊部"
    arrDepts(6) = "客服部"

    Dim arrCount(6) As Long
    arrCount(0) = 45 : arrCount(1) = 32 : arrCount(2) = 18
    arrCount(3) = 12 : arrCount(4) = 10 : arrCount(5) = 25
    arrCount(6) = 20

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
    objSheet.Cells(1, 1).Value = "部門"
    objSheet.Cells(1, 2).Value = "人數"

    With objSheet.Range("A1:B1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 6
        objSheet.Cells(i + 2, 1).Value = arrDepts(i)
        objSheet.Cells(i + 2, 2).Value = arrCount(i)
    Next i

    objSheet.Columns("A:B").AutoFit

    ' ── 插入群組橫條圖 ───────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(200, 20, 480, 300)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlBarClustered
    objChart.SetSourceData objSheet.Range("A1:B8")

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
