Attribute VB_Name = "PivotRadarChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotRadarChartExample
'功能說明: 建立樞紐分析雷達圖，展示各部門績效多維評估
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotRadarChart()
    Call CreatePivotRadarChart
End Sub

Sub CreatePivotRadarChart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim chtObj  As ChartObject
    Dim cht     As Chart

    On Error GoTo ErrHandler

    Set wsData  = RadarGetOrCreateWs("部門績效資料")
    Set wsPivot = RadarGetOrCreateWs("雷達圖樞紐")

    Call FillRadarData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="部門績效樞紐")

    With pt
        .PivotFields("評估項目").Orientation = xlRowField
        .PivotFields("評估項目").Position = 1
        .PivotFields("部門").Orientation = xlColumnField
        .PivotFields("部門").Position = 1
        .AddDataField .PivotFields("分數"), "平均分數", xlAverage
        .DataFields(1).NumberFormat = "0.0"
    End With

    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("F3").Left, _
        Top:=wsPivot.Range("F3").Top, _
        Width:=420, Height:=320)

    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlRadar
    cht.HasTitle = True
    cht.ChartTitle.Text = "各部門績效雷達圖"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    wsPivot.Activate
    MsgBox "雷達圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立雷達圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillRadarData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "評估項目"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "分數"
    ws.Range("A2").Value = "工作效率"
    ws.Range("B2").Value = "業務部"
    ws.Range("C2").Value = 85
    ws.Range("A3").Value = "團隊合作"
    ws.Range("B3").Value = "業務部"
    ws.Range("C3").Value = 78
    ws.Range("A4").Value = "創新能力"
    ws.Range("B4").Value = "業務部"
    ws.Range("C4").Value = 72
    ws.Range("A5").Value = "客戶服務"
    ws.Range("B5").Value = "業務部"
    ws.Range("C5").Value = 91
    ws.Range("A6").Value = "問題解決"
    ws.Range("B6").Value = "業務部"
    ws.Range("C6").Value = 80
    ws.Range("A7").Value = "工作效率"
    ws.Range("B7").Value = "技術部"
    ws.Range("C7").Value = 90
    ws.Range("A8").Value = "團隊合作"
    ws.Range("B8").Value = "技術部"
    ws.Range("C8").Value = 82
    ws.Range("A9").Value = "創新能力"
    ws.Range("B9").Value = "技術部"
    ws.Range("C9").Value = 95
    ws.Range("A10").Value = "客戶服務"
    ws.Range("B10").Value = "技術部"
    ws.Range("C10").Value = 70
    ws.Range("A11").Value = "問題解決"
    ws.Range("B11").Value = "技術部"
    ws.Range("C11").Value = 88
    ws.Columns("A:C").AutoFit
End Sub

Private Function RadarGetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set RadarGetOrCreateWs = ws
End Function
