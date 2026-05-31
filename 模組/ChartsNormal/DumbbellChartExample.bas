Attribute VB_Name = "DumbbellChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: DumbbellChartExample
'功能說明: 在Excel中建立啞鈴圖(Dumbbell Chart)範例，比較兩期間數值差距
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

' 測試進入口
Sub TestDumbbellChart()
    Call CreateDumbbellChart("啞鈴圖範例")
End Sub

' 建立啞鈴圖（以散點圖+連接線模擬）
Sub CreateDumbbellChart(ByVal sheetName As String)
    Dim ws         As Worksheet
    Dim chartObj   As ChartObject
    Dim cht        As Chart

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    Call FillDumbbellData(ws)

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=450, _
        Height:=320)

    Set cht = chartObj.Chart
    cht.ChartType = xlXYScatterLines
    cht.SetSourceData Source:=ws.Range("A1:C6")

    With cht.SeriesCollection(1)
        .Name = "今年"
        .XValues = ws.Range("B2:B6")
        .Values = ws.Range("A2:A6")
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 10
        .MarkerForegroundColor = RGB(0, 112, 192)
        .MarkerBackgroundColor = RGB(0, 112, 192)
    End With

    cht.SeriesCollection.NewSeries
    With cht.SeriesCollection(2)
        .Name = "去年"
        .XValues = ws.Range("C2:C6")
        .Values = ws.Range("A2:A6")
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 10
        .MarkerForegroundColor = RGB(255, 0, 0)
        .MarkerBackgroundColor = RGB(255, 0, 0)
    End With

    cht.HasTitle = True
    cht.ChartTitle.Text = "各部門今年 vs 去年業績比較"

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "部門"
    End With

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "業績（萬元）"
    End With

    cht.HasLegend = True

    MsgBox "啞鈴圖已建立完成！", vbInformation, "完成"
End Sub

' 填入啞鈴圖範例資料
Private Sub FillDumbbellData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "今年"
    ws.Range("C1").Value = "去年"
    ws.Range("A2").Value = 1
    ws.Range("B2").Value = 980
    ws.Range("C2").Value = 820
    ws.Range("A3").Value = 2
    ws.Range("B3").Value = 1150
    ws.Range("C3").Value = 1300
    ws.Range("A4").Value = 3
    ws.Range("B4").Value = 760
    ws.Range("C4").Value = 650
    ws.Range("A5").Value = 4
    ws.Range("B5").Value = 1420
    ws.Range("C5").Value = 1100
    ws.Range("A6").Value = 5
    ws.Range("B6").Value = 890
    ws.Range("C6").Value = 930
    ws.Columns("A:C").AutoFit
End Sub
