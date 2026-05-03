Attribute VB_Name = "CreateBubbleChart"
' ============================================================
' CreateBubbleChart.bas
' 說明：使用 Excel VBA 自動建立泡泡圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（城市 GDP / 人口 / 幸福指數）
'   3. 插入泡泡圖（X=GDP, Y=人口, 泡泡大小=幸福指數）
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateBubbleChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  As String = "城市 GDP、人口與幸福指數"
Const X_AXIS_TITLE As String = "GDP（千億元）"
Const Y_AXIS_TITLE As String = "人口（萬人）"
Const SHEET_NAME   As String = "城市資料"
Const OUTPUT_FILE  As String = "BubbleChartExample.xlsx"

Sub CreateBubbleChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    ' 欄位：城市, GDP（千億元）, 人口（萬人）, 幸福指數（泡泡大小）
    Dim arrCity(5) As String
    arrCity(0) = "台北" : arrCity(1) = "新北" : arrCity(2) = "桃園"
    arrCity(3) = "台中" : arrCity(4) = "台南" : arrCity(5) = "高雄"

    Dim arrGDP(5) As Long
    arrGDP(0) = 38 : arrGDP(1) = 22 : arrGDP(2) = 18
    arrGDP(3) = 20 : arrGDP(4) = 12 : arrGDP(5) = 16

    Dim arrPop(5) As Long
    arrPop(0) = 267 : arrPop(1) = 403 : arrPop(2) = 229
    arrPop(3) = 281 : arrPop(4) = 188 : arrPop(5) = 276

    Dim arrHappy(5) As Long
    arrHappy(0) = 75 : arrHappy(1) = 70 : arrHappy(2) = 72
    arrHappy(3) = 78 : arrHappy(4) = 80 : arrHappy(5) = 74

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook As Workbook
    Dim objSheet    As Worksheet
    Dim objChartObj As ChartObject
    Dim objChart    As Chart
    Dim objSeries   As Series
    Dim savePath    As String
    Dim i           As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook = Workbooks.Add()
    Set objSheet    = objWorkbook.Sheets(1)
    objSheet.Name   = SHEET_NAME

    ' ── 寫入標題列 ──────────────────────────────────────────────
    objSheet.Cells(1, 1).Value = "城市"
    objSheet.Cells(1, 2).Value = "GDP（千億元）"
    objSheet.Cells(1, 3).Value = "人口（萬人）"
    objSheet.Cells(1, 4).Value = "幸福指數"

    With objSheet.Range("A1:D1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 5
        objSheet.Cells(i + 2, 1).Value = arrCity(i)
        objSheet.Cells(i + 2, 2).Value = arrGDP(i)
        objSheet.Cells(i + 2, 3).Value = arrPop(i)
        objSheet.Cells(i + 2, 4).Value = arrHappy(i)
    Next i

    objSheet.Columns("A:D").AutoFit

    ' ── 插入泡泡圖 ───────────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(270, 20, 480, 300)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlBubble

    ' 泡泡圖需手動指定 X、Y 與泡泡大小數列
    objChart.SeriesCollection.NewSeries
    Set objSeries = objChart.SeriesCollection(1)
    objSeries.Name        = "城市"
    objSeries.XValues     = objSheet.Range("B2:B7")
    objSeries.Values      = objSheet.Range("C2:C7")
    objSeries.BubbleSizes = objSheet.Range("D2:D7")

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
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
    End With

    objChart.HasLegend = False

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objSeries   = Nothing
    Set objChart    = Nothing
    Set objChartObj = Nothing
    Set objSheet    = Nothing
    Set objWorkbook = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
