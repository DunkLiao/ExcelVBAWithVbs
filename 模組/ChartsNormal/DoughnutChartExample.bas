Attribute VB_Name = "DoughnutChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: DoughnutChartExample
'功能說明: 在Excel中建立圓環圖的範例程式
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口
Sub TestDoughnutChart()
    Call CreateDoughnutChart("圓環圖範例")
End Sub

' 建立圓環圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateDoughnutChart(ByVal sheetName As String)
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
    Call FillDepartmentData(ws)

    Set dataRange = ws.Range("A1:B6")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=400, _
        Height:=300)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlDoughnut
    cht.HasTitle = True
    cht.ChartTitle.Text = "各部門費用佔比"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight

    With cht.SeriesCollection(1)
        .HasDataLabels = True
        .DataLabels.ShowPercentage = True
        .DataLabels.ShowValue = False
        .DataLabels.Position = xlLabelPositionCenter
    End With

    cht.SeriesCollection(1).DoughnutHoleSize = 50

    MsgBox "圓環圖已建立完成！", vbInformation, "完成"
End Sub

' 填入部門費用範例資料
Private Sub FillDepartmentData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "費用"
    ws.Range("A2").Value = "研發部"
    ws.Range("B2").Value = 350000
    ws.Range("A3").Value = "業務部"
    ws.Range("B3").Value = 280000
    ws.Range("A4").Value = "行政部"
    ws.Range("B4").Value = 120000
    ws.Range("A5").Value = "人資部"
    ws.Range("B5").Value = 90000
    ws.Range("A6").Value = "財務部"
    ws.Range("B6").Value = 160000
    ws.Columns("A:B").AutoFit
End Sub
