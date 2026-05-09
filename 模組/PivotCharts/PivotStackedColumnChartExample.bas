Attribute VB_Name = "PivotStackedColumnChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotStackedColumnChartExample
'功能說明: 建立樞紐分析堆疊直條圖，展示各季度產品銷售組成
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotStackedColumnChart()
    Call CreatePivotStackedColumnChart
End Sub

Sub CreatePivotStackedColumnChart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim chtObj  As ChartObject
    Dim cht     As Chart

    On Error GoTo ErrHandler

    Set wsData  = StkColGetOrCreateWs("季度銷售資料")
    Set wsPivot = StkColGetOrCreateWs("堆疊直條樞紐")

    Call FillStackedColumnData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="季度銷售樞紐")

    With pt
        .PivotFields("季度").Orientation = xlRowField
        .PivotFields("季度").Position = 1
        .PivotFields("產品").Orientation = xlColumnField
        .PivotFields("產品").Position = 1
        .AddDataField .PivotFields("銷售額"), "銷售額合計", xlSum
        .DataFields(1).NumberFormat = "#,##0"
    End With

    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("F3").Left, _
        Top:=wsPivot.Range("F3").Top, _
        Width:=450, Height:=300)

    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlColumnStacked
    cht.HasTitle = True
    cht.ChartTitle.Text = "各季度產品銷售組成堆疊直條圖"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    wsPivot.Activate
    MsgBox "堆疊直條圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立堆疊直條圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillStackedColumnData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "季度"
    ws.Range("B1").Value = "產品"
    ws.Range("C1").Value = "銷售額"
    ws.Range("A2").Value = "Q1"
    ws.Range("B2").Value = "產品A"
    ws.Range("C2").Value = 150000
    ws.Range("A3").Value = "Q1"
    ws.Range("B3").Value = "產品B"
    ws.Range("C3").Value = 120000
    ws.Range("A4").Value = "Q1"
    ws.Range("B4").Value = "產品C"
    ws.Range("C4").Value = 90000
    ws.Range("A5").Value = "Q2"
    ws.Range("B5").Value = "產品A"
    ws.Range("C5").Value = 175000
    ws.Range("A6").Value = "Q2"
    ws.Range("B6").Value = "產品B"
    ws.Range("C6").Value = 135000
    ws.Range("A7").Value = "Q2"
    ws.Range("B7").Value = "產品C"
    ws.Range("C7").Value = 105000
    ws.Range("A8").Value = "Q3"
    ws.Range("B8").Value = "產品A"
    ws.Range("C8").Value = 200000
    ws.Range("A9").Value = "Q3"
    ws.Range("B9").Value = "產品B"
    ws.Range("C9").Value = 155000
    ws.Range("A10").Value = "Q3"
    ws.Range("B10").Value = "產品C"
    ws.Range("C10").Value = 120000
    ws.Range("A11").Value = "Q4"
    ws.Range("B11").Value = "產品A"
    ws.Range("C11").Value = 230000
    ws.Range("A12").Value = "Q4"
    ws.Range("B12").Value = "產品B"
    ws.Range("C12").Value = 180000
    ws.Range("A13").Value = "Q4"
    ws.Range("B13").Value = "產品C"
    ws.Range("C13").Value = 145000
    ws.Columns("A:C").AutoFit
End Sub

Private Function StkColGetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set StkColGetOrCreateWs = ws
End Function
