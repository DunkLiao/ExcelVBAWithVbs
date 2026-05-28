Attribute VB_Name = "PivotStepAreaChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotStepAreaChartExample
'功能說明: 從樞紐分析表建立面積圖，展示各期銷售額累計趨勢
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestPivotStepAreaChart()
    Call CreatePivotStepAreaChart(ThisWorkbook)
End Sub

Sub CreatePivotStepAreaChart(ByVal wb As Workbook)
    Dim dataWs    As Worksheet
    Dim pivotWs   As Worksheet
    Dim pc        As PivotCache
    Dim pt        As PivotTable
    Dim chartObj  As ChartObject
    Dim cht       As Chart

    On Error Resume Next
    Set dataWs = wb.Worksheets("階梯圖資料")
    On Error GoTo 0
    If dataWs Is Nothing Then
        Set dataWs = wb.Worksheets.Add
        dataWs.Name = "階梯圖資料"
    End If
    dataWs.Cells.Clear
    Call FillStepAreaData(dataWs)

    On Error Resume Next
    Set pivotWs = wb.Worksheets("階梯圖樞紐")
    On Error GoTo 0
    If pivotWs Is Nothing Then
        Set pivotWs = wb.Worksheets.Add(After:=dataWs)
        pivotWs.Name = "階梯圖樞紐"
    End If
    pivotWs.Cells.Clear

    On Error GoTo ErrHandler
    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataWs.UsedRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="階梯圖樞紐")

    With pt.PivotFields("季度")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields("銷售額")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0"
        .Name = "銷售合計"
    End With

    pivotWs.Columns.AutoFit

    Set chartObj = pivotWs.ChartObjects.Add( _
        Left:=pivotWs.Range("E1").Left, _
        Top:=pivotWs.Range("E1").Top, _
        Width:=420, _
        Height:=300)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=pivotWs.Range("A1").CurrentRegion
    cht.ChartType = xlArea
    cht.HasTitle = True
    cht.ChartTitle.Text = "各季度銷售額累計趨勢（面積圖）"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "季度"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售合計 (元)"
    End With

    cht.SeriesCollection(1).HasDataLabels = True
    cht.ChartStyle = 5
    pivotWs.Activate
    MsgBox "樞紐面積圖已建立完畢！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立圖表失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillStepAreaData(ByVal ws As Worksheet)
    ws.Range("A1:B1").Value = Array("季度", "銷售額")
    ws.Range("A2:B2").Value = Array("Q1", 185000)
    ws.Range("A3:B3").Value = Array("Q2", 243000)
    ws.Range("A4:B4").Value = Array("Q3", 318000)
    ws.Range("A5:B5").Value = Array("Q4", 412000)
    ws.Range("A6:B6").Value = Array("Q5", 295000)
    ws.Range("A7:B7").Value = Array("Q7", 428000)
    ws.Range("A8:B8").Value = Array("Q8", 510000)
    ws.Columns("A:B").AutoFit
End Sub
