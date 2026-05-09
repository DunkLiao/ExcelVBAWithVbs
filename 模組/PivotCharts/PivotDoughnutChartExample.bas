Attribute VB_Name = "PivotDoughnutChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotDoughnutChartExample
'功能說明: 建立樞紐分析環形圖，展示各產品類別市佔率
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotDoughnutChart()
    Call CreatePivotDoughnutChart
End Sub

Sub CreatePivotDoughnutChart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim chtObj  As ChartObject
    Dim cht     As Chart

    On Error GoTo ErrHandler

    Set wsData  = DoughnutGetOrCreateWs("產品市佔資料")
    Set wsPivot = DoughnutGetOrCreateWs("環形圖樞紐")

    Call FillDoughnutData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="產品市佔樞紐")

    With pt
        .PivotFields("產品類別").Orientation = xlRowField
        .PivotFields("產品類別").Position = 1
        .AddDataField .PivotFields("銷售額"), "銷售額合計", xlSum
        .DataFields(1).NumberFormat = "#,##0"
    End With

    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("D3").Left, _
        Top:=wsPivot.Range("D3").Top, _
        Width:=420, Height:=300)

    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlDoughnut
    cht.HasTitle = True
    cht.ChartTitle.Text = "各產品類別市佔率環形圖"
    cht.SeriesCollection(1).ApplyDataLabels
    cht.SeriesCollection(1).DataLabels.ShowPercentage = True
    cht.SeriesCollection(1).DataLabels.ShowValue = False
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight

    wsPivot.Activate
    MsgBox "環形圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立環形圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillDoughnutData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "產品類別"
    ws.Range("B1").Value = "銷售額"
    ws.Range("A2").Value = "電子產品"
    ws.Range("B2").Value = 350000
    ws.Range("A3").Value = "服飾配件"
    ws.Range("B3").Value = 210000
    ws.Range("A4").Value = "食品飲料"
    ws.Range("B4").Value = 180000
    ws.Range("A5").Value = "家居用品"
    ws.Range("B5").Value = 145000
    ws.Range("A6").Value = "運動休閒"
    ws.Range("B6").Value = 125000
    ws.Range("A7").Value = "書籍文具"
    ws.Range("B7").Value = 90000
    ws.Columns("A:B").AutoFit
End Sub

Private Function DoughnutGetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set DoughnutGetOrCreateWs = ws
End Function
