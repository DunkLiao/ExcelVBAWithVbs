Attribute VB_Name = "PivotSurfaceChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: 樞紐曲面圖範例
'功能說明: 以樞紐分析表資料建立曲面圖(Surface Chart)的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub TestPivotSurfaceChart()
    Call CreatePivotSurfaceChart("樞紐曲面圖範例")
End Sub

Sub CreatePivotSurfaceChart(ByVal sheetName As String)
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim cht As Chart

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    ThisWorkbook.Worksheets(sheetName & "_資料").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsData = ThisWorkbook.Worksheets.Add
    wsData.Name = sheetName & "_資料"
    Call FillSurfaceData(wsData)

    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = sheetName

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="PT_Surface")

    With pt
        .PivotFields("溫度").Orientation = xlRowField
        .PivotFields("溫度").Position = 1
        .PivotFields("壓力").Orientation = xlColumnField
        .PivotFields("壓力").Position = 1
        .AddDataField .PivotFields("產量"), "產量平均", xlAverage
    End With

    Set chartObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("H3").Left, _
        Top:=wsPivot.Range("H3").Top, _
        Width:=420, _
        Height:=320)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlSurface
    cht.HasTitle = True
    cht.ChartTitle.Text = "溫度與壓力對產量的影響"
    cht.HasLegend = True

    MsgBox "樞紐曲面圖已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillSurfaceData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "溫度"
    ws.Range("B1").Value = "壓力"
    ws.Range("C1").Value = "產量"

    ws.Range("A2").Value = 100 : ws.Range("B2").Value = 1 : ws.Range("C2").Value = 50
    ws.Range("A3").Value = 100 : ws.Range("B3").Value = 2 : ws.Range("C3").Value = 65
    ws.Range("A4").Value = 100 : ws.Range("B4").Value = 3 : ws.Range("C4").Value = 70
    ws.Range("A5").Value = 150 : ws.Range("B5").Value = 1 : ws.Range("C5").Value = 80
    ws.Range("A6").Value = 150 : ws.Range("B6").Value = 2 : ws.Range("C6").Value = 90
    ws.Range("A7").Value = 150 : ws.Range("B7").Value = 3 : ws.Range("C7").Value = 85
    ws.Range("A8").Value = 200 : ws.Range("B8").Value = 1 : ws.Range("C8").Value = 70
    ws.Range("A9").Value = 200 : ws.Range("B9").Value = 2 : ws.Range("C9").Value = 75
    ws.Range("A10").Value = 200 : ws.Range("B10").Value = 3 : ws.Range("C10").Value = 68

    ws.Columns("A:C").AutoFit
End Sub
