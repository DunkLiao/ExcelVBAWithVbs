Attribute VB_Name = "ColumnChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: ColumnChartExample
'功能說明: 在Excel中建立直條圖的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

' 範例入口
Sub TestColumnChart()
    Call CreateColumnChart("直條圖範例")
End Sub

' 建立單系列直條圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateColumnChart(ByVal sheetName As String)
    Dim ws       As Worksheet
    Dim chartObj As ChartObject
    Dim cht      As Chart
    Dim dataRng  As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    Call FillColumnData(ws)

    Set dataRng = ws.Range("A1:B7")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=400, _
        Height:=300)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRng
    cht.ChartType = xlColumnClustered

    cht.HasTitle = True
    cht.ChartTitle.Text = "各地區季度銷售量"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "地區"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售量"
    End With

    cht.ChartStyle = 5
    cht.SeriesCollection(1).HasDataLabels = True

    MsgBox "直條圖已建立完成！", vbInformation, "完成"
End Sub

' 填入範例資料
Private Sub FillColumnData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "銷售量"

    ws.Range("A2").Value = "北部"
    ws.Range("B2").Value = 2300
    ws.Range("A3").Value = "中部"
    ws.Range("B3").Value = 1850
    ws.Range("A4").Value = "南部"
    ws.Range("B4").Value = 2100
    ws.Range("A5").Value = "東部"
    ws.Range("B5").Value = 980
    ws.Range("A6").Value = "離島"
    ws.Range("B6").Value = 540
    ws.Range("A7").Value = "海外"
    ws.Range("B7").Value = 3200

    ws.Range("A1:B1").Font.Bold = True
    ws.Columns("A:B").AutoFit
End Sub
