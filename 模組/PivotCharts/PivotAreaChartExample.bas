Attribute VB_Name = "PivotAreaChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotAreaChartExample
'功能說明: 建立樞紐分析區域圖，展示各月份銷售趨勢
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotAreaChart()
    Call CreatePivotAreaChart
End Sub

Sub CreatePivotAreaChart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim chtObj  As ChartObject
    Dim cht     As Chart

    On Error GoTo ErrHandler

    Set wsData  = AreaGetOrCreateWs("月份銷售資料")
    Set wsPivot = AreaGetOrCreateWs("區域圖樞紐")

    Call FillMonthlyAreaData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="月份銷售樞紐")

    With pt
        .PivotFields("月份").Orientation = xlRowField
        .PivotFields("月份").Position = 1
        .AddDataField .PivotFields("銷售額"), "銷售額合計", xlSum
    End With

    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("E3").Left, _
        Top:=wsPivot.Range("E3").Top, _
        Width:=450, Height:=300)

    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlArea
    cht.HasTitle = True
    cht.ChartTitle.Text = "各月份銷售趨勢區域圖"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    wsPivot.Activate
    MsgBox "區域圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立區域圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillMonthlyAreaData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "銷售額"
    ws.Range("A2").Value = "一月"
    ws.Range("B2").Value = 85000
    ws.Range("A3").Value = "二月"
    ws.Range("B3").Value = 92000
    ws.Range("A4").Value = "三月"
    ws.Range("B4").Value = 110000
    ws.Range("A5").Value = "四月"
    ws.Range("B5").Value = 103000
    ws.Range("A6").Value = "五月"
    ws.Range("B6").Value = 128000
    ws.Range("A7").Value = "六月"
    ws.Range("B7").Value = 145000
    ws.Columns("A:B").AutoFit
End Sub

Private Function AreaGetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set AreaGetOrCreateWs = ws
End Function
