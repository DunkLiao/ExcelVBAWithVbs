Option Explicit

' 建立堆疊直條圖範例，呈現各月份不同區域的銷售占比。
Public Sub CreateStackedColumnChartExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim dataRange As Range
    Dim sheetName As String

    sheetName = "堆疊直條圖範例"
    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    Call FillStackedColumnSampleData(ws)
    Set dataRange = ws.Range("A1:D7")

    Set chartObj = ws.ChartObjects.Add(Left:=ws.Range("F2").Left, Top:=ws.Range("F2").Top, Width:=480, Height:=320)
    With chartObj.Chart
        .SetSourceData Source:=dataRange
        .ChartType = xlColumnStacked
        .HasTitle = True
        .ChartTitle.Text = "各區每月銷售堆疊圖"
        .HasLegend = True
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "月份"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "銷售金額"
    End With

    ws.Columns("A:D").AutoFit
    MsgBox "堆疊直條圖已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立堆疊直條圖失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillStackedColumnSampleData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("月份", "北區", "中區", "南區")
    ws.Range("A2:D2").Value = Array("一月", 120, 95, 88)
    ws.Range("A3:D3").Value = Array("二月", 135, 110, 92)
    ws.Range("A4:D4").Value = Array("三月", 148, 125, 105)
    ws.Range("A5:D5").Value = Array("四月", 142, 132, 118)
    ws.Range("A6:D6").Value = Array("五月", 160, 140, 126)
    ws.Range("A7:D7").Value = Array("六月", 175, 152, 139)
End Sub

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateWorksheet.Name = sheetName
    End If
End Function