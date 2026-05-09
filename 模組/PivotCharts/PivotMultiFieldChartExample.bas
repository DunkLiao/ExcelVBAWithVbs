Attribute VB_Name = "PivotMultiFieldChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotMultiFieldChartExample
'功能說明: 建立多欄位樞紐分析圖，進行地區與產品線交叉分析
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotMultiFieldChart()
    Call CreatePivotMultiFieldChart
End Sub

Sub CreatePivotMultiFieldChart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim chtObj  As ChartObject
    Dim cht     As Chart

    On Error GoTo ErrHandler

    Set wsData  = MultiGetOrCreateWs("地區產品資料")
    Set wsPivot = MultiGetOrCreateWs("多欄位樞紐")

    Call FillMultiFieldData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="地區產品樞紐")

    With pt
        .PivotFields("地區").Orientation = xlRowField
        .PivotFields("地區").Position = 1
        .PivotFields("產品線").Orientation = xlColumnField
        .PivotFields("產品線").Position = 1
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
    cht.ChartTitle.Text = "各地區產品線銷售量交叉分析圖"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom
    cht.ChartStyle = 10

    wsPivot.Activate
    MsgBox "多欄位樞紐圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立多欄位樞紐圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillMultiFieldData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "產品線"
    ws.Range("C1").Value = "銷售量"
    ws.Range("A2").Value = "北部"
    ws.Range("B2").Value = "家電"
    ws.Range("C2").Value = 520
    ws.Range("A3").Value = "北部"
    ws.Range("B3").Value = "3C"
    ws.Range("C3").Value = 880
    ws.Range("A4").Value = "北部"
    ws.Range("B4").Value = "服飾"
    ws.Range("C4").Value = 430
    ws.Range("A5").Value = "中部"
    ws.Range("B5").Value = "家電"
    ws.Range("C5").Value = 380
    ws.Range("A6").Value = "中部"
    ws.Range("B6").Value = "3C"
    ws.Range("C6").Value = 620
    ws.Range("A7").Value = "中部"
    ws.Range("B7").Value = "服飾"
    ws.Range("C7").Value = 350
    ws.Range("A8").Value = "南部"
    ws.Range("B8").Value = "家電"
    ws.Range("C8").Value = 410
    ws.Range("A9").Value = "南部"
    ws.Range("B9").Value = "3C"
    ws.Range("C9").Value = 550
    ws.Range("A10").Value = "南部"
    ws.Range("B10").Value = "服飾"
    ws.Range("C10").Value = 480
    ws.Columns("A:C").AutoFit
End Sub

Private Function MultiGetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set MultiGetOrCreateWs = ws
End Function
