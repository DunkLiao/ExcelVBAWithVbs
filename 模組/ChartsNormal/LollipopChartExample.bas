Attribute VB_Name = "LollipopChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: LollipopChartExample
'功能說明: 在Excel中建立棒棒糖圖表（Lollipop Chart）的示範程式，以散佈圖加誤差線呈現
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestLollipopChart()
    Call CreateLollipopChart("棒棒糖圖表")
End Sub

' 建立棒棒糖圖表
' sheetName: 要建立圖表的工作表名稱
Sub CreateLollipopChart(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range
    Dim ser As Series
    
    ' 取得或建立工作表
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillLollipopData(ws)
    
    ' 先建立散佈圖
    Set dataRange = ws.Range("B2:C8")
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=480, _
        Height:=320)
    
    Set cht = chartObj.Chart
    cht.ChartType = xlXYScatter
    cht.SetSourceData Source:=dataRange
    
    ' 設定X軸標籤
    cht.SeriesCollection(1).XValues = ws.Range("C2:C8")
    cht.SeriesCollection(1).Values = ws.Range("B2:B8")
    
    ' 加入垂直誤差線模擬棒棒糖的線條
    Set ser = cht.SeriesCollection(1)
    ser.HasErrorBars = True
    With ser.ErrorBars(xlY)
        .EndStyle = xlNoCap
        .Direction = xlY
        .Include = xlMinusValues
        .Type = xlFixedValue
        .Value = cht.Axes(xlValue).MaximumScale
    End With
    
    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "產品銷售棒棒糖圖"
    
    ' 設定軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "銷售額"
    End With
    
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "產品名稱"
    End With
    
    ' 設定樣式
    cht.ChartStyle = 2
    ser.MarkerSize = 10
    ser.MarkerStyle = xlMarkerStyleCircle
    
    MsgBox "棒棒糖圖表已建立完成！", vbInformation, "完成"
End Sub

' 填入棒棒糖圖表示範資料
Private Sub FillLollipopData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "銷售額"
    ws.Range("C1").Value = "X軸位置"
    
    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 850
    ws.Range("C2").Value = 1
    
    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 620
    ws.Range("C3").Value = 2
    
    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 430
    ws.Range("C4").Value = 3
    
    ws.Range("A5").Value = "產品D"
    ws.Range("B5").Value = 780
    ws.Range("C5").Value = 4
    
    ws.Range("A6").Value = "產品E"
    ws.Range("B6").Value = 550
    ws.Range("C6").Value = 5
    
    ws.Range("A7").Value = "產品F"
    ws.Range("B7").Value = 920
    ws.Range("C7").Value = 6
    
    ws.Range("A8").Value = "產品G"
    ws.Range("B8").Value = 680
    ws.Range("C8").Value = 7
    
    ws.Columns("A:C").AutoFit
End Sub
