Attribute VB_Name = "PivotHistogramChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotHistogramChartExample
'功能說明: 以樞紐分析表為資料來源，建立直方圖（Histogram）樞紐分析圖範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub CreatePivotHistogramChart()
    Dim ws As Worksheet
    Dim ptWs As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim scores As Variant
    Dim i As Integer

    Const DATA_SHEET As String = "直方圖原始資料"
    Const PIVOT_SHEET As String = "直方圖樞紐"

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = DATA_SHEET
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1").Value = "學生"
    ws.Range("B1").Value = "成績"

    scores = Array(55, 62, 70, 72, 75, 75, 78, 80, 80, 82, 83, 85, 85, 87, 88, 90, 92, 93, 95, 98)

    For i = 0 To UBound(scores)
        ws.Cells(i + 2, 1).Value = "學生" & (i + 1)
        ws.Cells(i + 2, 2).Value = scores(i)
    Next i

    On Error Resume Next
    Set ptWs = ThisWorkbook.Worksheets(PIVOT_SHEET)
    On Error GoTo 0

    Application.DisplayAlerts = False
    If Not ptWs Is Nothing Then ptWs.Delete
    Application.DisplayAlerts = True

    Set ptWs = ThisWorkbook.Worksheets.Add(After:=ws)
    ptWs.Name = PIVOT_SHEET

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ws.Range("A1:B" & (UBound(scores) + 2)))

    Set pt = pc.CreatePivotTable( _
        TableDestination:=ptWs.Range("A3"), _
        TableName:="直方圖樞紐")

    With pt
        .PivotFields("成績").Orientation = xlRowField
        .PivotFields("成績").Position = 1
        .AddDataField .PivotFields("成績"), "計數", xlCount
    End With

    Set chartObj = ptWs.ChartObjects.Add( _
        Left:=ptWs.Range("E1").Left, _
        Top:=ptWs.Range("E1").Top, _
        Width:=420, _
        Height:=300)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlColumnClustered

    cht.HasTitle = True
    cht.ChartTitle.Text = "成績分佈直方圖"

    MsgBox "直方圖樞紐分析圖已建立完成！", vbInformation, "完成"
End Sub
