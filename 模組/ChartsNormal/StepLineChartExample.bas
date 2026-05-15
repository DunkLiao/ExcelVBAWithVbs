Attribute VB_Name = "StepLineChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: StepLineChartExample
'功能說明: 階梯折線圖範例，建立月份與數值資料並產生階梯折線圖
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunStepLineChartExample()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim stepRange As Range

    Set ws = GetOrCreateStepSheet("階梯折線圖範例")
    ws.Cells.Clear

    Call FillStepSourceData(ws)
    Set stepRange = BuildStepChartData(ws)
    Call RemoveAllChartObjects(ws)
    Call RemoveStepAnnotations(ws)

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("G2").Left, _
        Top:=ws.Range("G2").Top, _
        Width:=480, _
        Height:=300)

    With chartObj.Chart
        .ChartType = xlLine
        .SetSourceData Source:=stepRange
        .HasTitle = True
        .ChartTitle.Text = "每月值階梯折線圖"
        .HasLegend = False
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "月份"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "數值"
        .SeriesCollection(1).Smooth = False
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).Format.Line.Weight = 2.25
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(1).MarkerSize = 6
    End With

    Call AddStepAnnotations(ws, chartObj)
    ws.Columns("A:E").AutoFit

    MsgBox "階梯折線圖已建立完成。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立階梯折線圖時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillStepSourceData(ByVal ws As Worksheet)
    ws.Range("A1:B1").Value = Array("月份", "數值")
    ws.Range("A2:B2").Value = Array("1月", 80)
    ws.Range("A3:B3").Value = Array("2月", 80)
    ws.Range("A4:B4").Value = Array("3月", 125)
    ws.Range("A5:B5").Value = Array("4月", 125)
    ws.Range("A6:B6").Value = Array("5月", 160)
    ws.Range("A7:B7").Value = Array("6月", 145)
End Sub

Private Function BuildStepChartData(ByVal ws As Worksheet) As Range
    Dim lastRow As Long
    Dim i As Long
    Dim stepRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Range("D:E").Clear
    ws.Range("D1:E1").Value = Array("階梯月份", "階梯值")

    stepRow = 2
    ws.Cells(stepRow, 4).Value = ws.Cells(2, 1).Value
    ws.Cells(stepRow, 5).Value = ws.Cells(2, 2).Value
    stepRow = stepRow + 1

    For i = 3 To lastRow
        ws.Cells(stepRow, 4).Value = ws.Cells(i, 1).Value & " 前"
        ws.Cells(stepRow, 5).Value = ws.Cells(i - 1, 2).Value
        stepRow = stepRow + 1

        ws.Cells(stepRow, 4).Value = ws.Cells(i, 1).Value
        ws.Cells(stepRow, 5).Value = ws.Cells(i, 2).Value
        stepRow = stepRow + 1
    Next i

    Set BuildStepChartData = ws.Range(ws.Cells(1, 4), ws.Cells(stepRow - 1, 5))
End Function

Private Sub AddStepAnnotations(ByVal ws As Worksheet, ByVal chartObj As ChartObject)
    Dim lastSourceRow As Long
    Dim latestValue As Variant
    Dim maxValue As Double
    Dim maxMonth As String
    Dim i As Long

    lastSourceRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    latestValue = ws.Cells(lastSourceRow, 2).Value
    maxValue = ws.Cells(2, 2).Value
    maxMonth = CStr(ws.Cells(2, 1).Value)

    For i = 3 To lastSourceRow
        If CDbl(ws.Cells(i, 2).Value) > maxValue Then
            maxValue = CDbl(ws.Cells(i, 2).Value)
            maxMonth = CStr(ws.Cells(i, 1).Value)
        End If
    Next i

    With ws.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=chartObj.Left + 12, _
        Top:=chartObj.Top + 12, _
        Width:=150, _
        Height:=36)
        .Name = "StepAnnotationLatest"
        .TextFrame.Characters.Text = "最新數值: " & latestValue
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(255, 242, 204)
    End With

    With ws.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=chartObj.Left + 12, _
        Top:=chartObj.Top + 52, _
        Width:=170, _
        Height:=36)
        .Name = "StepAnnotationPeak"
        .TextFrame.Characters.Text = "最高月份: " & maxMonth & " / " & maxValue
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(226, 239, 218)
    End With
End Sub

Private Sub RemoveAllChartObjects(ByVal ws As Worksheet)
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
End Sub

Private Sub RemoveStepAnnotations(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Shapes("StepAnnotationLatest").Delete
    ws.Shapes("StepAnnotationPeak").Delete
    On Error GoTo 0
End Sub

Private Function GetOrCreateStepSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateStepSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateStepSheet Is Nothing Then
        Set GetOrCreateStepSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateStepSheet.Name = sheetName
    End If
End Function
