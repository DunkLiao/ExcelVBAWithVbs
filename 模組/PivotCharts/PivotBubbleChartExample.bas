Attribute VB_Name = "PivotBubbleChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotBubbleChartExample
'功能說明: 以樞紐彙總資料建立泡泡圖，展示市場分析的三維資料視覺呈現
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口
Sub TestPivotBubbleChart()
    Call CreatePivotBubbleChart("泡泡圖範例")
End Sub

' 建立樞紐泡泡圖範例
' sheetName: 工作表名稱
Sub CreatePivotBubbleChart(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim ser As Series

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Call FillBubblePivotData(ws)

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("F1").Left, _
        Top:=ws.Range("F1").Top, _
        Width:=480, _
        Height:=360)

    Set cht = chartObj.Chart
    cht.ChartType = xlBubble

    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    Set ser = cht.SeriesCollection.NewSeries
    With ser
        .Name = "產品A"
        .XValues = ws.Range("B3:B5")
        .Values = ws.Range("C3:C5")
        .BubbleSizes = ws.Range("D3:D5")
    End With

    Set ser = cht.SeriesCollection.NewSeries
    With ser
        .Name = "產品B"
        .XValues = ws.Range("B7:B9")
        .Values = ws.Range("C7:C9")
        .BubbleSizes = ws.Range("D7:D9")
    End With

    cht.HasTitle = True
    cht.ChartTitle.Text = "市場分析泡泡圖（市占率 vs 成長率）"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "市場占有率 (%)"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "成長率 (%)"
    End With

    cht.HasLegend = True
    ws.Columns("A:E").AutoFit
    MsgBox "泡泡圖已建立完成！", vbInformation, "完成"
End Sub

' 填入泡泡圖樞紐彙總資料
Private Sub FillBubblePivotData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "樞紐彙總：市場分析"
    ws.Range("A1").Font.Bold = True

    ws.Range("A2").Value = "產品線"
    ws.Range("B2").Value = "市占率(%)"
    ws.Range("C2").Value = "成長率(%)"
    ws.Range("D2").Value = "銷售量"
    ws.Range("A2:D2").Font.Bold = True

    ws.Range("A3").Value = "A-高階": ws.Range("B3").Value = 25: ws.Range("C3").Value = 15: ws.Range("D3").Value = 300
    ws.Range("A4").Value = "A-中階": ws.Range("B4").Value = 18: ws.Range("C4").Value = 22: ws.Range("D4").Value = 500
    ws.Range("A5").Value = "A-低階": ws.Range("B5").Value = 10: ws.Range("C5").Value = 8:  ws.Range("D5").Value = 800
    ws.Range("A6").Value = ""
    ws.Range("A7").Value = "B-高階": ws.Range("B7").Value = 30: ws.Range("C7").Value = 5:  ws.Range("D7").Value = 250
    ws.Range("A8").Value = "B-中階": ws.Range("B8").Value = 22: ws.Range("C8").Value = 18: ws.Range("D8").Value = 420
    ws.Range("A9").Value = "B-低階": ws.Range("B9").Value = 8:  ws.Range("C9").Value = 30: ws.Range("D9").Value = 650
End Sub