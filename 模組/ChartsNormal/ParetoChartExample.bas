Option Explicit
Attribute VB_Name = "ParetoChartExample"
'*************************************************************************************
'模組名稱: ParetoChartExample
'功能說明: 建立帕雷托圖（Pareto Chart），以長條圖顯示各原因數量，並疊加累計百分比折線
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub CreateParetoChartExample()
    Dim ws As Worksheet
    Dim rngData As Range
    Dim cht As ChartObject
    Dim objChart As Chart
    Dim i As Integer
    Dim total As Double
    Dim cumPct As Double
    Dim srs As Series
    Dim axisY2 As Axis

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear

    ws.Range("A1").Value = "原因類別"
    ws.Range("B1").Value = "發生次數"
    ws.Range("C1").Value = "累計百分比"

    Dim categories(1 To 6) As String
    Dim counts(1 To 6) As Long
    categories(1) = "設備故障" : counts(1) = 120
    categories(2) = "操作失誤" : counts(2) = 95
    categories(3) = "材料不良" : counts(3) = 60
    categories(4) = "環境因素" : counts(4) = 30
    categories(5) = "設計缺陷" : counts(5) = 20
    categories(6) = "其他原因" : counts(6) = 10

    For i = 1 To 6
        ws.Cells(i + 1, 1).Value = categories(i)
        ws.Cells(i + 1, 2).Value = counts(i)
    Next i

    total = 0
    For i = 1 To 6
        total = total + counts(i)
    Next i

    cumPct = 0
    For i = 1 To 6
        cumPct = cumPct + counts(i)
        ws.Cells(i + 1, 3).Value = Round(cumPct / total * 100, 1)
    Next i

    ws.Range("C2:C7").NumberFormat = "0.0"
    ws.Columns("A:C").AutoFit

    Set cht = ws.ChartObjects.Add(Left:=10, Top:=130, Width:=480, Height:=280)
    Set objChart = cht.Chart

    Set rngData = ws.Range("A1:B7")
    objChart.SetSourceData Source:=rngData
    objChart.ChartType = xlColumnClustered

    Set srs = objChart.SeriesCollection.NewSeries
    srs.Values = ws.Range("C2:C7")
    srs.Name = "累計百分比"
    srs.ChartType = xlLine
    srs.AxisGroup = xlSecondary

    objChart.HasTitle = True
    objChart.ChartTitle.Text = "帕雷托圖範例"

    Set axisY2 = objChart.Axes(xlValue, xlSecondary)
    axisY2.MaximumScale = 100
    axisY2.MinimumScale = 0

    MsgBox "帕雷托圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "建立帕雷托圖失敗"
End Sub