Attribute VB_Name = "CompareWithChartVisual"
Option Explicit
'*************************************************************************************
'模組名稱: CompareWithChartVisual
'功能說明: 自動比較兩組資料差異，並以圖表視覺化呈現比較結果的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestCompareWithChartVisual()
    Call CompareWithChartVisual
End Sub

' 比較兩組資料並以圖表呈現差異
Sub CompareWithChartVisual()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim lastRow As Long
    Dim i As Long
    
    sheetName = "圖表比較差異"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillCompareData(ws)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 計算差異值與差異百分比
    ws.Range("D1").Value = "差異值"
    ws.Range("E1").Value = "差異百分比"
    
    For i = 2 To lastRow
        ws.Cells(i, 4).Formula = "=C" & i & "-B" & i
        ws.Cells(i, 5).Formula = "=IF(B" & i & "=0,0,(C" & i & "-B" & i & ")/B" & i & ")"
        ws.Cells(i, 5).NumberFormat = "0.0%"
    Next i
    
    ' 建立比較圖表
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("G1").Left, _
        Top:=ws.Range("G1").Top, _
        Width:=500, _
        Height:=320)
    
    Set cht = chartObj.Chart
    cht.ChartType = xlColumnClustered
    
    ' 加入去年資料系列
    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(1).Name = "去年"
    cht.SeriesCollection(1).Values = ws.Range("B2:B" & lastRow)
    cht.SeriesCollection(1).XValues = ws.Range("A2:A" & lastRow)
    
    ' 加入今年資料系列
    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(2).Name = "今年"
    cht.SeriesCollection(2).Values = ws.Range("C2:C" & lastRow)
    cht.SeriesCollection(2).XValues = ws.Range("A2:A" & lastRow)
    
    ' 加入差異值折線
    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(3).Name = "差異值"
    cht.SeriesCollection(3).Values = ws.Range("D2:D" & lastRow)
    cht.SeriesCollection(3).ChartType = xlLineMarkers
    cht.SeriesCollection(3).XValues = ws.Range("A2:A" & lastRow)
    
    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "年度銷售比較與差異分析"
    
    ' 設定軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "月份"
    End With
    
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售額"
    End With
    
    ' 設定圖例
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom
    
    ' 樣式設定
    cht.ChartStyle = 10
    
    ' 用條件式格式標示差異
    With ws.Range("D2:D" & lastRow)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(200, 255, 200)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(255, 200, 200)
    End With
    
    ws.Columns("A:E").AutoFit
    ws.Activate
    
    MsgBox "資料差異比較與圖表視覺化已完成！", vbInformation, "完成"
End Sub

' 填入比較示範資料
Private Sub FillCompareData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "去年銷售額"
    ws.Range("C1").Value = "今年銷售額"
    
    ws.Range("A2").Value = "1月"
    ws.Range("B2").Value = 1200
    ws.Range("C2").Value = 1350
    
    ws.Range("A3").Value = "2月"
    ws.Range("B3").Value = 1100
    ws.Range("C3").Value = 1080
    
    ws.Range("A4").Value = "3月"
    ws.Range("B4").Value = 1400
    ws.Range("C4").Value = 1550
    
    ws.Range("A5").Value = "4月"
    ws.Range("B5").Value = 1300
    ws.Range("C5").Value = 1420
    
    ws.Range("A6").Value = "5月"
    ws.Range("B6").Value = 1500
    ws.Range("C6").Value = 1600
    
    ws.Range("A7").Value = "6月"
    ws.Range("B7").Value = 1600
    ws.Range("C7").Value = 1480
    
    ws.Columns("A:C").AutoFit
End Sub
