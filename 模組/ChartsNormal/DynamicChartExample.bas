Option Explicit
Attribute VB_Name = "DynamicChartExample"
'*************************************************************************************
'模組名稱: DynamicChartExample
'功能說明: 建立動態折線圖，示範即時切換圖表資料來源的功能
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestDynamicChart()
    Call CreateDynamicChart("動態圖表範例")
End Sub

Sub CreateDynamicChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart

    Set ws = GetOrCreateDynSheet(sheetName)
    ws.Cells.Clear

    Call FillDynamicChartData(ws)

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("G1").Left, _
        Top:=ws.Range("G1").Top, _
        Width:=480, _
        Height:=300)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=ws.Range("A1:D7")
    cht.ChartType = xlLine

    cht.HasTitle = True
    cht.ChartTitle.Text = "各月份業績趨勢（動態更新）"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "月份"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "業績（萬元）"
    End With

    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    MsgBox "動態折線圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立動態圖表時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillDynamicChartData(ByVal ws As Worksheet)
    Randomize

    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "產品A"
    ws.Range("C1").Value = "產品B"
    ws.Range("D1").Value = "產品C"

    Dim monthNames As Variant
    monthNames = Array("一月", "二月", "三月", "四月", "五月", "六月")

    Dim r As Integer
    For r = 2 To 7
        ws.Cells(r, 1).Value = monthNames(r - 2)
        ws.Cells(r, 2).Value = Int(Rnd() * 100) + 100
        ws.Cells(r, 3).Value = Int(Rnd() * 100) + 80
        ws.Cells(r, 4).Value = Int(Rnd() * 100) + 60
    Next r

    ws.Columns("A:D").AutoFit
End Sub

Private Function GetOrCreateDynSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateDynSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateDynSheet Is Nothing Then
        Set GetOrCreateDynSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateDynSheet.Name = sheetName
    End If
End Function
