Attribute VB_Name = "BandChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: 帶狀圖範例
'功能說明: 在Excel中建立帶狀折線圖(Band Chart)的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub TestBandChart()
    Call CreateBandChart("帶狀圖範例")
End Sub

Sub CreateBandChart(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    Call FillBandData(ws)

    Set dataRange = ws.Range("A1:D7")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("F1").Left, _
        Top:=ws.Range("F1").Top, _
        Width:=420, _
        Height:=300)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlLine

    ' 將下限系列設定為透明面積填滿，呈現帶狀效果
    With cht.SeriesCollection(2)
        .ChartType = xlArea
        .Format.Fill.ForeColor.RGB = RGB(198, 224, 180)
        .Format.Fill.Transparency = 0.5
        .Border.LineStyle = xlNone
    End With

    With cht.SeriesCollection(3)
        .ChartType = xlArea
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Fill.Transparency = 0
        .Border.LineStyle = xlNone
    End With

    cht.HasTitle = True
    cht.ChartTitle.Text = "每月銷售額與帶狀目標區間"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "月份"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售額"
    End With

    cht.HasLegend = True

    MsgBox "帶狀圖已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillBandData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "實際銷售額"
    ws.Range("C1").Value = "目標下限"
    ws.Range("D1").Value = "目標上限"

    ws.Range("A2").Value = "一月" : ws.Range("B2").Value = 820 : ws.Range("C2").Value = 750 : ws.Range("D2").Value = 900
    ws.Range("A3").Value = "二月" : ws.Range("B3").Value = 910 : ws.Range("C3").Value = 800 : ws.Range("D3").Value = 950
    ws.Range("A4").Value = "三月" : ws.Range("B4").Value = 760 : ws.Range("C4").Value = 780 : ws.Range("D4").Value = 920
    ws.Range("A5").Value = "四月" : ws.Range("B5").Value = 980 : ws.Range("C5").Value = 820 : ws.Range("D5").Value = 970
    ws.Range("A6").Value = "五月" : ws.Range("B6").Value = 860 : ws.Range("C6").Value = 800 : ws.Range("D6").Value = 960
    ws.Range("A7").Value = "六月" : ws.Range("B7").Value = 1020 : ws.Range("C7").Value = 850 : ws.Range("D7").Value = 1000

    ws.Columns("A:D").AutoFit
End Sub
