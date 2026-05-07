Attribute VB_Name = "HistogramChartExample"
Option Explicit

' ============================================================
' 範例：建立長條圖（直方圖）並設定格式
' 功能：在新工作表上建立頻率分布直方圖
' ============================================================
Sub CreateHistogramChartExample()
    Dim ws      As Worksheet
    Dim chObj   As ChartObject
    Dim ch      As Chart
    Dim rngData As Range

    ' --- 準備示範資料 ---
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "HistogramDemo"

    ws.Cells(1, 1).Value = "區間"
    ws.Cells(1, 2).Value = "頻率"
    ws.Cells(2, 1).Value = "0-10"
    ws.Cells(2, 2).Value = 5
    ws.Cells(3, 1).Value = "11-20"
    ws.Cells(3, 2).Value = 12
    ws.Cells(4, 1).Value = "21-30"
    ws.Cells(4, 2).Value = 20
    ws.Cells(5, 1).Value = "31-40"
    ws.Cells(5, 2).Value = 18
    ws.Cells(6, 1).Value = "41-50"
    ws.Cells(6, 2).Value = 8

    Set rngData = ws.Range("A1:B6")

    ' --- 建立圖表物件 ---
    Set chObj = ws.ChartObjects.Add(Left:=20, Top:=80, Width:=400, Height:=260)
    Set ch = chObj.Chart

    ch.SetSourceData Source:=rngData, PlotBy:=xlColumns
    ch.ChartType = xlColumnClustered

    ' --- 設定標題與軸標籤 ---
    ch.HasTitle = True
    ch.ChartTitle.Text = "頻率分布直方圖"

    ch.Axes(xlCategory, xlPrimary).HasTitle = True
    ch.Axes(xlCategory, xlPrimary).AxisTitle.Text = "區間"

    ch.Axes(xlValue, xlPrimary).HasTitle = True
    ch.Axes(xlValue, xlPrimary).AxisTitle.Text = "頻率"

    ' --- 移除數列間隙，使直方圖外觀更貼近 ---
    ch.SeriesCollection(1).GapWidth = 0
    ch.SeriesCollection(1).Interior.Color = RGB(70, 130, 180)

    MsgBox "直方圖已建立於工作表：" & ws.Name, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
