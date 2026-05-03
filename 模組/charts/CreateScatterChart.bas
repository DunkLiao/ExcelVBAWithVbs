Attribute VB_Name = "CreateScatterChart"
' ============================================================
' CreateScatterChart.bas
' 說明：使用 Excel VBA 自動建立散佈圖（XY圖）範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（身高與體重的關係）
'   3. 插入散佈圖
'   4. 設定圖表標題、座標軸標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateScatterChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  As String = "身高與體重分布"
Const X_AXIS_TITLE As String = "身高（cm）"
Const Y_AXIS_TITLE As String = "體重（kg）"
Const SHEET_NAME   As String = "身高體重資料"
Const OUTPUT_FILE  As String = "ScatterChartExample.xlsx"

Sub CreateScatterChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    ' 身高（cm）
    Dim arrHeight(9) As Long
    arrHeight(0) = 158 : arrHeight(1) = 162 : arrHeight(2) = 165
    arrHeight(3) = 168 : arrHeight(4) = 170 : arrHeight(5) = 172
    arrHeight(6) = 175 : arrHeight(7) = 178 : arrHeight(8) = 180
    arrHeight(9) = 185

    ' 對應體重（kg）
    Dim arrWeight(9) As Long
    arrWeight(0) = 52 : arrWeight(1) = 56 : arrWeight(2) = 58
    arrWeight(3) = 62 : arrWeight(4) = 65 : arrWeight(5) = 68
    arrWeight(6) = 70 : arrWeight(7) = 74 : arrWeight(8) = 77
    arrWeight(9) = 82

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
    objSheet.Cells(1, 1).Value = "身高（cm）"
    objSheet.Cells(1, 2).Value = "體重（kg）"

    With objSheet.Range("A1:B1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 9
        objSheet.Cells(i + 2, 1).Value = arrHeight(i)
        objSheet.Cells(i + 2, 2).Value = arrWeight(i)
    Next i

    objSheet.Columns("A:B").AutoFit

    ' ── 插入散佈圖 ───────────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(200, 20, 480, 300)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlXYScatter
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
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
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
