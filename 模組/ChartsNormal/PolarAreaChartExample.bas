Attribute VB_Name = "PolarAreaChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PolarAreaChartExample
'功能說明: 以VBA在Excel中建立極座標面積圖（以雷達填色圖模擬）範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestPolarAreaChart()
    Call CreatePolarAreaChart("極座標面積圖範例")
End Sub

' 建立極座標面積圖（以雷達填色圖模擬）
Sub CreatePolarAreaChart(ByVal sheetName As String)
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
    Call FillPolarData(ws)

    Set dataRange = ws.Range("A1:B9")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=400, _
        Height:=350)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlRadarFilled

    cht.HasTitle = True
    cht.ChartTitle.Text = "各方向風速分布（極座標面積圖）"

    cht.HasLegend = False
    cht.SeriesCollection(1).HasDataLabels = True

    ws.Columns("A:B").AutoFit
    ws.Activate
    MsgBox "極座標面積圖已建立完成！", vbInformation, "完成"
End Sub

' 填入極座標範例資料
Private Sub FillPolarData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "方向"
    ws.Range("B1").Value = "風速(km/h)"
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A2").Value = "北"
    ws.Range("B2").Value = 25
    ws.Range("A3").Value = "東北"
    ws.Range("B3").Value = 18
    ws.Range("A4").Value = "東"
    ws.Range("B4").Value = 32
    ws.Range("A5").Value = "東南"
    ws.Range("B5").Value = 15
    ws.Range("A6").Value = "南"
    ws.Range("B6").Value = 28
    ws.Range("A7").Value = "西南"
    ws.Range("B7").Value = 20
    ws.Range("A8").Value = "西"
    ws.Range("B8").Value = 35
    ws.Range("A9").Value = "西北"
    ws.Range("B9").Value = 22
    ws.Columns("A:B").AutoFit
End Sub
