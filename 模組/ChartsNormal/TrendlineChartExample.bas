Attribute VB_Name = "TrendlineChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: TrendlineChartExample
'功能說明: 建立含趨勢線的散點圖範例，示範如何在圖表中加入線性趨勢線
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

' 簡化呼叫入口
Sub TestTrendlineChart()
    Call CreateTrendlineChart("趨勢線圖範例")
End Sub

' 建立含趨勢線的散點圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateTrendlineChart(ByVal sheetName As String)
    Dim ws         As Worksheet
    Dim chartObj   As ChartObject
    Dim cht        As Chart
    Dim ser        As Series
    Dim dataRange  As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    ' 填入範例資料
    Call FillTrendlineData(ws)

    ' 設定資料範圍 (A1:B11)
    Set dataRange = ws.Range("A1:B11")

    ' 在工作表新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=440, _
        Height:=320)

    Set cht = chartObj.Chart

    ' 設定資料來源為散點圖
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlXYScatter

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "銷售量與廣告費用散點趨勢分析"

    ' 設定座標軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "廣告費用 (萬元)"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售量 (件)"
    End With

    ' 取得數列並加入趨勢線
    Set ser = cht.SeriesCollection(1)

    ser.Trendlines.Add( _
        Type:=xlLinear, _
        Forward:=0, _
        Backward:=0, _
        DisplayEquation:=True, _
        DisplayRSquared:=True, _
        Name:="線性趨勢")

    ' 設定趨勢線樣式
    With ser.Trendlines(1).Border
        .Color = RGB(255, 0, 0)
        .Weight = xlMedium
    End With

    ' 設定圖表樣式
    cht.ChartStyle = 7

    MsgBox "含趨勢線的散點圖已建立完畢！", vbInformation, "完成"
End Sub

' 填入趨勢線圖範例資料
Private Sub FillTrendlineData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "廣告費用"
    ws.Range("B1").Value = "銷售量"
    ws.Range("A2").Value = 2  : ws.Range("B2").Value = 120
    ws.Range("A3").Value = 4  : ws.Range("B3").Value = 180
    ws.Range("A4").Value = 5  : ws.Range("B4").Value = 210
    ws.Range("A5").Value = 7  : ws.Range("B5").Value = 260
    ws.Range("A6").Value = 8  : ws.Range("B6").Value = 310
    ws.Range("A7").Value = 10 : ws.Range("B7").Value = 350
    ws.Range("A8").Value = 12 : ws.Range("B8").Value = 390
    ws.Range("A9").Value = 14 : ws.Range("B9").Value = 420
    ws.Range("A10").Value = 16 : ws.Range("B10").Value = 480
    ws.Range("A11").Value = 18 : ws.Range("B11").Value = 510
    ws.Columns("A:B").AutoFit
End Sub
