Attribute VB_Name = "PivotForecastChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotForecastChartExample
'功能說明: 在樞紐分析圖中加入趨勢線與預測的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestPivotForecastChart()
    Call CreatePivotChartWithForecast
End Sub

Sub CreatePivotChartWithForecast()
    Dim wsData As Worksheet
    Dim wsChart As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim cht As Chart

    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsData = ThisWorkbook.Worksheets("預測圖表資料")
    If Not wsData Is Nothing Then wsData.Delete
    Set wsChart = ThisWorkbook.Worksheets("預測樞紐圖表")
    If Not wsChart Is Nothing Then wsChart.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 建立資料
    Set wsData = ThisWorkbook.Worksheets.Add
    wsData.Name = "預測圖表資料"

    wsData.Range("A1").Value = "月份"
    wsData.Range("B1").Value = "銷售額"
    wsData.Range("A1:B1").Font.Bold = True

    Dim months As Variant
    months = Array("1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月")

    Dim sales As Variant
    sales = Array(1200, 1350, 1500, 1400, 1600, 1750, 1800, 1950)

    Dim i As Integer
    For i = 0 To 7
        wsData.Cells(i + 2, 1).Value = months(i)
        wsData.Cells(i + 2, 2).Value = sales(i)
    Next i

    wsData.Columns("A:B").AutoFit

    ' 建立樞紐分析圖
    Set wsChart = ThisWorkbook.Worksheets.Add
    wsChart.Name = "預測樞紐圖表"

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1:B9"))

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsChart.Range("A1"), _
        TableName:="預測樞紐")

    pt.AddFields RowFields:="月份"

    With pt.PivotFields("銷售額")
        .Orientation = xlDataField
        .Function = xlSum
    End With

    ' 建立圖表
    Set cht = wsChart.Shapes.AddChart2( _
        Style:=201, _
        XlChartType:=xlLineMarkers).Chart

    cht.SetSourceData Source:=pt.TableRange1

    cht.HasTitle = True
    cht.ChartTitle.Text = "銷售趨勢與預測"

    ' 加入線性趨勢線
    cht.SeriesCollection(1).Trendlines.Add( _
        Type:=xlLinear, _
        Forward:=3, _
        DisplayEquation:=True, _
        DisplayRSquared:=True)

    MsgBox "包含趨勢預測線的樞紐分析圖已建立完成！", vbInformation, "完成"
End Sub
