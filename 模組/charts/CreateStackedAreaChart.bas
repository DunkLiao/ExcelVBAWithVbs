Attribute VB_Name = "CreateStackedAreaChart"
' ============================================================
' CreateStackedAreaChart.bas
' 說明：使用 Excel VBA 自動建立堆疊區域圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（月度電力來源組成）
'   3. 插入堆疊區域圖
'   4. 設定圖表標題、座標軸標籤、圖例等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateStackedAreaChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE  As String = "2025 年月度電力來源組成（億度）"
Const X_AXIS_TITLE As String = "月份"
Const Y_AXIS_TITLE As String = "發電量（億度）"
Const SHEET_NAME   As String = "電力來源"
Const OUTPUT_FILE  As String = "StackedAreaChartExample.xlsx"

Sub CreateStackedAreaChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    Dim arrMonths(11) As String
    arrMonths(0)  = "1月"  : arrMonths(1)  = "2月"  : arrMonths(2)  = "3月"
    arrMonths(3)  = "4月"  : arrMonths(4)  = "5月"  : arrMonths(5)  = "6月"
    arrMonths(6)  = "7月"  : arrMonths(7)  = "8月"  : arrMonths(8)  = "9月"
    arrMonths(9)  = "10月" : arrMonths(10) = "11月" : arrMonths(11) = "12月"

    ' 火力發電（億度）
    Dim arrFire(11) As Long
    arrFire(0)  = 120 : arrFire(1)  = 110 : arrFire(2)  = 105
    arrFire(3)  = 100 : arrFire(4)  = 115 : arrFire(5)  = 130
    arrFire(6)  = 140 : arrFire(7)  = 135 : arrFire(8)  = 120
    arrFire(9)  = 110 : arrFire(10) = 115 : arrFire(11) = 125

    ' 核能發電（億度）
    Dim arrNuclear(11) As Long
    arrNuclear(0)  = 40 : arrNuclear(1)  = 38 : arrNuclear(2)  = 42
    arrNuclear(3)  = 41 : arrNuclear(4)  = 40 : arrNuclear(5)  = 43
    arrNuclear(6)  = 45 : arrNuclear(7)  = 44 : arrNuclear(8)  = 42
    arrNuclear(9)  = 40 : arrNuclear(10) = 39 : arrNuclear(11) = 41

    ' 再生能源（億度）
    Dim arrRenew(11) As Long
    arrRenew(0)  = 20 : arrRenew(1)  = 22 : arrRenew(2)  = 28
    arrRenew(3)  = 30 : arrRenew(4)  = 32 : arrRenew(5)  = 35
    arrRenew(6)  = 38 : arrRenew(7)  = 36 : arrRenew(8)  = 30
    arrRenew(9)  = 25 : arrRenew(10) = 22 : arrRenew(11) = 20

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
    objSheet.Cells(1, 2).Value = "火力發電"
    objSheet.Cells(1, 3).Value = "核能發電"
    objSheet.Cells(1, 4).Value = "再生能源"

    With objSheet.Range("A1:D1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 11
        objSheet.Cells(i + 2, 1).Value = arrMonths(i)
        objSheet.Cells(i + 2, 2).Value = arrFire(i)
        objSheet.Cells(i + 2, 3).Value = arrNuclear(i)
        objSheet.Cells(i + 2, 4).Value = arrRenew(i)
    Next i

    objSheet.Columns("A:D").AutoFit

    ' ── 插入堆疊區域圖 ───────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(260, 20, 480, 300)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlAreaStacked
    objChart.SetSourceData objSheet.Range("A1:D13")

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
