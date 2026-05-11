Attribute VB_Name = "BoxWhiskerChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: BoxWhiskerChartExample
'功能說明: 在Excel中建立箱型圖（Box and Whisker）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub TestBoxWhiskerChart()
    Call CreateBoxWhiskerChart("箱型圖範例")
End Sub

Sub CreateBoxWhiskerChart(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    Call FillBoxWhiskerData(ws)

    Set dataRange = ws.Range("A1:C11")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=400, _
        Height:=300)

    Set cht = chartObj.Chart

    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlBoxwhisker

    cht.HasTitle = True
    cht.ChartTitle.Text = "各組成績分佈箱型圖"

    cht.HasLegend = True

    MsgBox "箱型圖已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillBoxWhiskerData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "組別"
    ws.Range("B1").Value = "甲班"
    ws.Range("C1").Value = "乙班"

    Dim scores1 As Variant
    Dim scores2 As Variant
    scores1 = Array(55, 62, 70, 75, 78, 80, 83, 88, 92, 95)
    scores2 = Array(48, 58, 65, 70, 74, 79, 85, 89, 91, 98)

    Dim i As Integer
    For i = 1 To 10
        ws.Range("A" & (i + 1)).Value = "數據" & i
        ws.Range("B" & (i + 1)).Value = scores1(i - 1)
        ws.Range("C" & (i + 1)).Value = scores2(i - 1)
    Next i

    ws.Columns("A:C").AutoFit
End Sub
