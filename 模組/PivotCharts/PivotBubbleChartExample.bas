Attribute VB_Name = "PivotBubbleChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotBubbleChartExample
'功能說明: 建立樞紐分析表氣泡圖的示範程式，結合樞紐分析與氣泡圖呈現三維度資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestPivotBubbleChart()
    Call CreatePivotBubbleChart
End Sub

' 建立樞紐氣泡圖
Sub CreatePivotBubbleChart()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim dataRange As Range
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim lastRow As Long
    
    ' 建立資料工作表
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("氣泡圖資料")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Worksheets.Add
        wsData.Name = "氣泡圖資料"
    End If
    
    wsData.Cells.Clear
    Call FillBubbleChartData(wsData)
    
    ' 建立樞紐分析表
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("氣泡圖樞紐")
    If Not wsPivot Is Nothing Then wsPivot.Delete
    On Error GoTo 0
    
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "氣泡圖樞紐"
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set dataRange = wsData.Range("A1:D" & lastRow)
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange, _
        Version:=xlPivotTableVersion15)
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="PivotBubble")
    
    With pt
        .PivotFields("產品").Orientation = xlRowField
        .PivotFields("產品").Position = 1
    End With
    
    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum
    pt.AddDataField pt.PivotFields("利潤"), "利潤合計", xlSum
    pt.AddDataField pt.PivotFields("市佔率"), "市佔率平均", xlAverage
    
    ' 取得樞紐分析表範圍
    Dim pivotRange As Range
    Dim rowCount As Long
    
    Set pivotRange = pt.TableRange1
    rowCount = pt.RowRange.Rows.Count
    
    ' 建立樞紐圖表（氣泡圖）
    Set chartObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("A1").Left, _
        Top:=wsPivot.Range("A1").Top + pivotRange.Height + 20, _
        Width:=520, _
        Height:=360)
    
    Set cht = chartObj.Chart
    cht.ChartType = xlBubble
    
    ' 使用樞紐結果繪製氣泡圖
    If rowCount > 1 Then
        ' 清除預設數列
        Do While cht.SeriesCollection.Count > 0
            cht.SeriesCollection(1).Delete
        Loop
        
        ' 建立氣泡圖數列
        cht.SeriesCollection.NewSeries
        cht.SeriesCollection(1).Name = "產品分析"
        cht.SeriesCollection(1).XValues = wsPivot.Range("B2:B" & rowCount)
        cht.SeriesCollection(1).Values = wsPivot.Range("C2:C" & rowCount)
        cht.SeriesCollection(1).BubbleSizes = "=" & wsPivot.Name & "!" & "D2:D" & rowCount
    End If
    
    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "樞紐分析氣泡圖 - 產品銷售分析"
    
    ' 設定軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "銷售額"
    End With
    
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "利潤"
    End With
    
    ' 設定圖例
    cht.HasLegend = True
    
    ' 套用樣式
    cht.ChartStyle = 8
    
    wsPivot.Activate
    
    MsgBox "樞紐分析氣泡圖已建立完成！" & vbCrLf & _
           "氣泡大小代表市佔率，X軸為銷售額，Y軸為利潤。", vbInformation, "完成"
End Sub

' 填入氣泡圖示範資料
Private Sub FillBubbleChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "銷售額"
    ws.Range("C1").Value = "利潤"
    ws.Range("D1").Value = "市佔率"
    
    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 85000
    ws.Range("C2").Value = 25000
    ws.Range("D2").Value = 0.15
    
    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 62000
    ws.Range("C3").Value = 18000
    ws.Range("D3").Value = 0.08
    
    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 140000
    ws.Range("C4").Value = 45000
    ws.Range("D4").Value = 0.35
    
    ws.Range("A5").Value = "產品D"
    ws.Range("B5").Value = 45000
    ws.Range("C5").Value = 10000
    ws.Range("D5").Value = 0.05
    
    ws.Range("A6").Value = "產品E"
    ws.Range("B6").Value = 110000
    ws.Range("C6").Value = 38000
    ws.Range("D6").Value = 0.25
    
    ws.Range("A7").Value = "產品F"
    ws.Range("B7").Value = 32000
    ws.Range("C7").Value = 8000
    ws.Range("D7").Value = 0.03
    
    ws.Columns("A:D").AutoFit
End Sub
