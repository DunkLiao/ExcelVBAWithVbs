Attribute VB_Name = "ErrorBarChartExample"
Option Explicit
'*************************************************************************************
'家舱嘿: ErrorBarChartExample
'弧: Excelいミ粇畉次ч絬瓜絛ㄒ祘Α
'
'舦┮Τ: Dunk
'祘Α砞璸: Dunk
'级糶ら戳: 2026/5/13
'
'*************************************************************************************

' 璶
Sub TestErrorBarChart()
    Call CreateErrorBarChart("粇畉次瓜絛ㄒ")
End Sub

' ミ粇畉次ч絬瓜
Sub CreateErrorBarChart(ByVal sheetName As String)
    Dim ws        As Worksheet
    Dim chartObj  As ChartObject
    Dim cht       As Chart
    Dim ser       As Series
    Dim dataRange As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    Call FillErrorBarData(ws)

    Set dataRange = ws.Range("A1:B7")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=420, _
        Height:=300)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlLine
    cht.HasTitle = True
    cht.ChartTitle.Text = "–るキА放粇畉次"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "る"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "放C"
    End With

    Set ser = cht.SeriesCollection(1)
    ser.HasErrorBars = True
    With ser.ErrorBar(xlY, xlBoth, xlFixedValue, 2)
    End With
    cht.ChartStyle = 4
    ser.HasDataLabels = False
    MsgBox "粇畉次ч絬瓜ミЧΘ", vbInformation, "ЧΘ"
End Sub

' 恶絛ㄒ放计沮
Private Sub FillErrorBarData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "る"
    ws.Range("B1").Value = "キА放"
    ws.Range("A2").Value = "る" : ws.Range("B2").Value = 10
    ws.Range("A3").Value = "る" : ws.Range("B3").Value = 12
    ws.Range("A4").Value = "る" : ws.Range("B4").Value = 18
    ws.Range("A5").Value = "る" : ws.Range("B5").Value = 24
    ws.Range("A6").Value = "きる" : ws.Range("B6").Value = 29
    ws.Range("A7").Value = "せる" : ws.Range("B7").Value = 33
    ws.Columns("A:B").AutoFit
End Sub

