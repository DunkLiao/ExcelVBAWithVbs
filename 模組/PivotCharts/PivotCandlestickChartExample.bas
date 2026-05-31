Attribute VB_Name = "PivotCandlestickChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotCandlestickChartExample
'功能說明: 以樞紐分析結果建立K線（股價）圖，顯示開高低收四個數值
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestPivotCandlestickChart()
    Call CreatePivotCandlestickChart("K線圖範例")
End Sub

Sub CreatePivotCandlestickChart(ByVal sheetName As String)
    Dim ws       As Worksheet
    Dim chartObj As ChartObject
    Dim cht      As Chart

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    Call FillCandlestickData(ws)

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("G1").Left, _
        Top:=ws.Range("G1").Top, _
        Width:=480, _
        Height:=300)

    Set cht = chartObj.Chart
    cht.ChartType = xlStockOHLC
    cht.SetSourceData Source:=ws.Range("A1:E6")

    cht.HasTitle = True
    cht.ChartTitle.Text = "股票K線圖（開高低收）"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "日期"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "股價"
    End With

    cht.HasLegend = True

    MsgBox "K線圖已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillCandlestickData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "開盤"
    ws.Range("C1").Value = "最高"
    ws.Range("D1").Value = "最低"
    ws.Range("E1").Value = "收盤"

    ws.Range("A2").Value = "2026/5/25"
    ws.Range("B2").Value = 150
    ws.Range("C2").Value = 158
    ws.Range("D2").Value = 148
    ws.Range("E2").Value = 155

    ws.Range("A3").Value = "2026/5/26"
    ws.Range("B3").Value = 155
    ws.Range("C3").Value = 162
    ws.Range("D3").Value = 153
    ws.Range("E3").Value = 160

    ws.Range("A4").Value = "2026/5/27"
    ws.Range("B4").Value = 160
    ws.Range("C4").Value = 165
    ws.Range("D4").Value = 156
    ws.Range("E4").Value = 158

    ws.Range("A5").Value = "2026/5/28"
    ws.Range("B5").Value = 158
    ws.Range("C5").Value = 161
    ws.Range("D5").Value = 152
    ws.Range("E5").Value = 154

    ws.Range("A6").Value = "2026/5/29"
    ws.Range("B6").Value = 154
    ws.Range("C6").Value = 159
    ws.Range("D6").Value = 150
    ws.Range("E6").Value = 157

    ws.Range("A2:A6").NumberFormat = "yyyy/m/d"
    ws.Columns("A:E").AutoFit
End Sub
