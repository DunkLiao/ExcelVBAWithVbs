Attribute VB_Name = "PivotStackedBarChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotStackedBarChartExample
'功能說明: 建立樞紐分析堆疊橫條圖，分析各部門費用結構
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotStackedBarChart()
    Call CreatePivotStackedBarChart
End Sub

Sub CreatePivotStackedBarChart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim chtObj  As ChartObject
    Dim cht     As Chart

    On Error GoTo ErrHandler

    Set wsData  = StkBarGetOrCreateWs("部門費用資料")
    Set wsPivot = StkBarGetOrCreateWs("堆疊橫條樞紐")

    Call FillStackedBarData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="部門費用樞紐")

    With pt
        .PivotFields("部門").Orientation = xlRowField
        .PivotFields("部門").Position = 1
        .PivotFields("費用類型").Orientation = xlColumnField
        .PivotFields("費用類型").Position = 1
        .AddDataField .PivotFields("費用金額"), "費用合計", xlSum
        .DataFields(1).NumberFormat = "#,##0"
    End With

    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("F3").Left, _
        Top:=wsPivot.Range("F3").Top, _
        Width:=450, Height:=300)

    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlBarStacked
    cht.HasTitle = True
    cht.ChartTitle.Text = "各部門費用結構堆疊橫條圖"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    wsPivot.Activate
    MsgBox "堆疊橫條圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立堆疊橫條圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillStackedBarData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "費用類型"
    ws.Range("C1").Value = "費用金額"
    ws.Range("A2").Value = "業務部"
    ws.Range("B2").Value = "人事費用"
    ws.Range("C2").Value = 320000
    ws.Range("A3").Value = "業務部"
    ws.Range("B3").Value = "差旅費用"
    ws.Range("C3").Value = 85000
    ws.Range("A4").Value = "業務部"
    ws.Range("B4").Value = "行銷費用"
    ws.Range("C4").Value = 120000
    ws.Range("A5").Value = "技術部"
    ws.Range("B5").Value = "人事費用"
    ws.Range("C5").Value = 450000
    ws.Range("A6").Value = "技術部"
    ws.Range("B6").Value = "差旅費用"
    ws.Range("C6").Value = 30000
    ws.Range("A7").Value = "技術部"
    ws.Range("B7").Value = "行銷費用"
    ws.Range("C7").Value = 20000
    ws.Range("A8").Value = "管理部"
    ws.Range("B8").Value = "人事費用"
    ws.Range("C8").Value = 280000
    ws.Range("A9").Value = "管理部"
    ws.Range("B9").Value = "差旅費用"
    ws.Range("C9").Value = 45000
    ws.Range("A10").Value = "管理部"
    ws.Range("B10").Value = "行銷費用"
    ws.Range("C10").Value = 60000
    ws.Columns("A:C").AutoFit
End Sub

Private Function StkBarGetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set StkBarGetOrCreateWs = ws
End Function
