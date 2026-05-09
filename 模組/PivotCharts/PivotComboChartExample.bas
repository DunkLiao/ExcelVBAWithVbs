Attribute VB_Name = "PivotComboChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotComboChartExample
'功能說明: 建立樞紐分析組合圖（直條圖+折線圖），呈現銷售額與達成率對比
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotComboChart()
    Call CreatePivotComboChart
End Sub

Sub CreatePivotComboChart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim chtObj  As ChartObject
    Dim cht     As Chart
    Dim ser     As Series

    On Error GoTo ErrHandler

    Set wsData  = ComboGetOrCreateWs("銷售達成資料")
    Set wsPivot = ComboGetOrCreateWs("組合圖樞紐")

    Call FillComboData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="銷售達成樞紐")

    With pt
        .PivotFields("業務員").Orientation = xlRowField
        .PivotFields("業務員").Position = 1
        .AddDataField .PivotFields("銷售額"), "銷售額合計", xlSum
        .AddDataField .PivotFields("達成率"), "平均達成率", xlAverage
        .DataFields(1).NumberFormat = "#,##0"
        .DataFields(2).NumberFormat = "0.0%"
    End With

    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("F3").Left, _
        Top:=wsPivot.Range("F3").Top, _
        Width:=480, Height:=320)

    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlColumnClustered

    '將第二數列改為折線圖並置於副座標軸
    If cht.SeriesCollection.Count >= 2 Then
        Set ser = cht.SeriesCollection(2)
        ser.ChartType = xlLine
        ser.AxisGroup = xlSecondary
    End If

    cht.HasTitle = True
    cht.ChartTitle.Text = "業務員銷售額與達成率組合圖"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    wsPivot.Activate
    MsgBox "組合圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立組合圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillComboData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "業務員"
    ws.Range("B1").Value = "銷售額"
    ws.Range("C1").Value = "達成率"
    ws.Range("A2").Value = "陳大明"
    ws.Range("B2").Value = 285000
    ws.Range("C2").Value = 0.95
    ws.Range("A3").Value = "林小華"
    ws.Range("B3").Value = 320000
    ws.Range("C3").Value = 1.07
    ws.Range("A4").Value = "王志遠"
    ws.Range("B4").Value = 198000
    ws.Range("C4").Value = 0.66
    ws.Range("A5").Value = "張美玲"
    ws.Range("B5").Value = 410000
    ws.Range("C5").Value = 1.37
    ws.Range("A6").Value = "李俊賢"
    ws.Range("B6").Value = 250000
    ws.Range("C6").Value = 0.83
    ws.Columns("A:C").AutoFit
End Sub

Private Function ComboGetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set ComboGetOrCreateWs = ws
End Function
