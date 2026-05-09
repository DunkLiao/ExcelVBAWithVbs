Attribute VB_Name = "PivotColumnChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotColumnChartExample
'功能說明: 建立樞紐分析群組直條圖，比較各月份不同產品銷售量
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotColumnChart()
    Call CreatePivotColumnChart
End Sub

Sub CreatePivotColumnChart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim chtObj  As ChartObject
    Dim cht     As Chart

    On Error GoTo ErrHandler

    Set wsData  = ColGetOrCreateWs("月份產品銷售")
    Set wsPivot = ColGetOrCreateWs("群組直條樞紐")

    Call FillColumnData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="月份產品樞紐")

    With pt
        .PivotFields("月份").Orientation = xlRowField
        .PivotFields("月份").Position = 1
        .PivotFields("產品").Orientation = xlColumnField
        .PivotFields("產品").Position = 1
        .AddDataField .PivotFields("銷售量"), "銷售量合計", xlSum
    End With

    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("F3").Left, _
        Top:=wsPivot.Range("F3").Top, _
        Width:=480, Height:=320)

    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlColumnClustered
    cht.HasTitle = True
    cht.ChartTitle.Text = "各月份產品銷售量群組直條圖"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom
    cht.ChartStyle = 6
    cht.SeriesCollection(1).HasDataLabels = True

    wsPivot.Activate
    MsgBox "群組直條圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立群組直條圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillColumnData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "產品"
    ws.Range("C1").Value = "銷售量"
    ws.Range("A2").Value = "一月"
    ws.Range("B2").Value = "手機"
    ws.Range("C2").Value = 320
    ws.Range("A3").Value = "一月"
    ws.Range("B3").Value = "平板"
    ws.Range("C3").Value = 185
    ws.Range("A4").Value = "二月"
    ws.Range("B4").Value = "手機"
    ws.Range("C4").Value = 295
    ws.Range("A5").Value = "二月"
    ws.Range("B5").Value = "平板"
    ws.Range("C5").Value = 210
    ws.Range("A6").Value = "三月"
    ws.Range("B6").Value = "手機"
    ws.Range("C6").Value = 410
    ws.Range("A7").Value = "三月"
    ws.Range("B7").Value = "平板"
    ws.Range("C7").Value = 240
    ws.Range("A8").Value = "四月"
    ws.Range("B8").Value = "手機"
    ws.Range("C8").Value = 380
    ws.Range("A9").Value = "四月"
    ws.Range("B9").Value = "平板"
    ws.Range("C9").Value = 195
    ws.Columns("A:C").AutoFit
End Sub

Private Function ColGetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set ColGetOrCreateWs = ws
End Function
