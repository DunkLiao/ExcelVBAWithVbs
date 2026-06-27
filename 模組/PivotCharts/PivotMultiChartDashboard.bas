Attribute VB_Name = "PivotMultiChartDashboard"
Option Explicit
'*************************************************************************************
'模組名稱: PivotMultiChartDashboard
'功能說明: 以單一樞紐分析表為資料來源，建立多個樞紐圖表組成儀表板（Dashboard）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestPivotMultiChartDashboard()
    Call CreatePivotMultiChartDashboard
End Sub

Sub CreatePivotMultiChartDashboard()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim wsDash As Worksheet
    Dim dataRange As Range
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj1 As ChartObject
    Dim chartObj2 As ChartObject
    Dim chartObj3 As ChartObject
    Dim cht1 As Chart
    Dim cht2 As Chart
    Dim cht3 As Chart
    Dim lastRow As Long
    
    ' 建立資料工作表
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("儀表板資料")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Worksheets.Add
        wsData.Name = "儀表板資料"
    End If
    wsData.Cells.Clear
    Call FillDashboardData(wsData)
    
    ' 建立樞紐分析表
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("儀表板樞紐")
    If Not wsPivot Is Nothing Then wsPivot.Delete
    On Error GoTo 0
    
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "儀表板樞紐"
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set dataRange = wsData.Range("A1:D" & lastRow)
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange, _
        Version:=xlPivotTableVersion15)
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="DashPivot")
    
    With pt
        .PivotFields("區域").Orientation = xlRowField
        .PivotFields("區域").Position = 1
        .PivotFields("類別").Orientation = xlColumnField
        .PivotFields("類別").Position = 1
    End With
    
    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum
    pt.AddDataField pt.PivotFields("數量"), "數量合計", xlSum
    
    ' 建立儀表板工作表
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets("儀表板")
    If Not wsDash Is Nothing Then wsDash.Delete
    On Error GoTo 0
    
    Set wsDash = ThisWorkbook.Worksheets.Add
    wsDash.Name = "儀表板"
    
    ' 圖表1: 長條圖 - 區域銷售比較
    Set chartObj1 = wsDash.ChartObjects.Add( _
        Left:=10, Top:=10, Width:=400, Height:=250)
    Set cht1 = chartObj1.Chart
    cht1.ChartType = xlColumnClustered
    cht1.SetSourceData Source:=pt.TableRange1
    cht1.HasTitle = True
    cht1.ChartTitle.Text = "區域銷售比較"
    cht1.ChartStyle = 2
    
    ' 圖表2: 圓餅圖 - 區域佔比
    Set chartObj2 = wsDash.ChartObjects.Add( _
        Left:=10, Top:=270, Width:=400, Height:=250)
    Set cht2 = chartObj2.Chart
    cht2.ChartType = xlPie
    cht2.SetSourceData Source:=pt.TableRange1
    cht2.HasTitle = True
    cht2.ChartTitle.Text = "區域銷售佔比"
    cht2.ApplyDataLabels
    cht2.SeriesCollection(1).DataLabels.ShowPercentage = True
    
    ' 圖表3: 折線圖 - 數量趨勢
    Set chartObj3 = wsDash.ChartObjects.Add( _
        Left:=420, Top:=10, Width:=400, Height:=250)
    Set cht3 = chartObj3.Chart
    cht3.ChartType = xlLineMarkers
    
    If cht3.SeriesCollection.Count > 0 Then
        Dim j As Long
        For j = cht3.SeriesCollection.Count To 1 Step -1
            cht3.SeriesCollection(j).Delete
        Next j
    End If
    
    cht3.SeriesCollection.NewSeries
    cht3.SeriesCollection(1).Name = "數量"
    cht3.SeriesCollection(1).Values = pt.DataBodyRange.Columns(2)
    cht3.SeriesCollection(1).XValues = pt.RowRange
    
    cht3.HasTitle = True
    cht3.ChartTitle.Text = "銷售數量趨勢"
    cht3.ChartStyle = 10
    
    ' 標題
    wsDash.Range("A1").Value = "銷售儀表板"
    wsDash.Range("A1").Font.Size = 16
    wsDash.Range("A1").Font.Bold = True
    
    wsDash.Activate
    
    MsgBox "樞紐儀表板已建立完成！" & vbCrLf & _
           "包含三種圖表：長條圖、圓餅圖、折線圖。", vbInformation, "完成"
End Sub

Private Sub FillDashboardData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "區域"
    ws.Range("B1").Value = "類別"
    ws.Range("C1").Value = "銷售額"
    ws.Range("D1").Value = "數量"
    
    ws.Range("A2").Value = "北部"
    ws.Range("B2").Value = "電子"
    ws.Range("C2").Value = 5000
    ws.Range("D2").Value = 120
    ws.Range("A3").Value = "中部"
    ws.Range("B3").Value = "電子"
    ws.Range("C3").Value = 3500
    ws.Range("D3").Value = 85
    ws.Range("A4").Value = "南部"
    ws.Range("B4").Value = "電子"
    ws.Range("C4").Value = 2800
    ws.Range("D4").Value = 70
    ws.Range("A5").Value = "北部"
    ws.Range("B5").Value = "食品"
    ws.Range("C5").Value = 4200
    ws.Range("D5").Value = 200
    ws.Range("A6").Value = "中部"
    ws.Range("B6").Value = "食品"
    ws.Range("C6").Value = 3100
    ws.Range("D6").Value = 150
    ws.Range("A7").Value = "南部"
    ws.Range("B7").Value = "食品"
    ws.Range("C7").Value = 2600
    ws.Range("D7").Value = 130
    ws.Range("A8").Value = "北部"
    ws.Range("B8").Value = "服飾"
    ws.Range("C8").Value = 3800
    ws.Range("D8").Value = 90
    ws.Range("A9").Value = "中部"
    ws.Range("B9").Value = "服飾"
    ws.Range("C9").Value = 2900
    ws.Range("D9").Value = 75
    ws.Range("A10").Value = "南部"
    ws.Range("B10").Value = "服飾"
    ws.Range("C10").Value = 2200
    ws.Range("D10").Value = 60
    
    ws.Columns("A:D").AutoFit
End Sub
