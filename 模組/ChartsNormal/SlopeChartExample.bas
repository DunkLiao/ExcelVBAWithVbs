Attribute VB_Name = "SlopeChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: SlopeChartExample
'功能說明: 建立斜率圖（Slope Chart），以折線圖比較資料項目前後的變化趨勢
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestSlopeChart()
    Call CreateSlopeChart("斜率圖")
End Sub

Sub CreateSlopeChart(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillSlopeData(ws)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=500, _
        Height:=350)
    
    Set cht = chartObj.Chart
    cht.ChartType = xlLineMarkers
    
    For i = 2 To lastRow
        cht.SeriesCollection.NewSeries
        cht.SeriesCollection(i - 1).Name = ws.Cells(i, 1).Value
        cht.SeriesCollection(i - 1).Values = ws.Range("B" & i & ":C" & i)
        cht.SeriesCollection(i - 1).XValues = ws.Range("B1:C1")
        cht.SeriesCollection(i - 1).MarkerStyle = xlMarkerStyleCircle
        cht.SeriesCollection(i - 1).MarkerSize = 8
    Next i
    
    cht.HasTitle = True
    cht.ChartTitle.Text = "部門預算年度變化斜率圖"
    
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "年度"
    End With
    
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "預算金額（萬元）"
    End With
    
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.ChartStyle = 2
    
    ws.Activate
    MsgBox "斜率圖已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillSlopeData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "去年"
    ws.Range("C1").Value = "今年"
    
    ws.Range("A2").Value = "業務部"
    ws.Range("B2").Value = 450
    ws.Range("C2").Value = 520
    
    ws.Range("A3").Value = "研發部"
    ws.Range("B3").Value = 380
    ws.Range("C3").Value = 410
    
    ws.Range("A4").Value = "行銷部"
    ws.Range("B4").Value = 280
    ws.Range("C4").Value = 320
    
    ws.Range("A5").Value = "人事部"
    ws.Range("B5").Value = 120
    ws.Range("C5").Value = 130
    
    ws.Range("A6").Value = "財務部"
    ws.Range("B6").Value = 150
    ws.Range("C6").Value = 140
    
    ws.Range("A7").Value = "資訊部"
    ws.Range("B7").Value = 200
    ws.Range("C7").Value = 260
    
    ws.Columns("A:C").AutoFit
End Sub
