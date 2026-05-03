Option Explicit
Attribute VB_Name = "BubbleChartExample"
'*************************************************************************************
'模組名稱: 泡泡圖範例
'功能描述: 在 Excel 中建立泡泡圖的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/4
'
'傳入參數:
'傳回值:
'
'*************************************************************************************

' 測試用入口
Sub TestBubbleChart()
    Call CreateBubbleChart("泡泡圖範例")
End Sub

' 建立泡泡圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateBubbleChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim seriesItem As Series

    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    Call FillBubbleChartData(ws)

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("F1").Left, _
        Top:=ws.Range("F1").Top, _
        Width:=480, _
        Height:=340)

    Set cht = chartObj.Chart
    cht.ChartType = xlBubble
    cht.HasTitle = True
    cht.ChartTitle.Text = "廣告支出、銷售額與客戶數分析"
    cht.HasLegend = False

    Set seriesItem = cht.SeriesCollection.NewSeries
    With seriesItem
        .Name = "市場資料"
        .XValues = ws.Range("B2:B8")
        .Values = ws.Range("C2:C8")
        .BubbleSizes = ws.Range("D2:D8")
        .HasDataLabels = True
        .DataLabels.ShowSeriesName = False
        .DataLabels.ShowCategoryName = False
        .DataLabels.ShowValue = False
    End With

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "廣告支出（萬元）"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售額（萬元）"
    End With

    cht.ChartStyle = 18

    MsgBox "泡泡圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立泡泡圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateWorksheet.Name = sheetName
    End If
End Function

' 輸入泡泡圖範例資料
Private Sub FillBubbleChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "區域"
    ws.Range("B1").Value = "廣告支出"
    ws.Range("C1").Value = "銷售額"
    ws.Range("D1").Value = "客戶數"

    ws.Range("A2").Value = "北區"
    ws.Range("B2").Value = 35
    ws.Range("C2").Value = 260
    ws.Range("D2").Value = 120

    ws.Range("A3").Value = "中區"
    ws.Range("B3").Value = 28
    ws.Range("C3").Value = 210
    ws.Range("D3").Value = 95

    ws.Range("A4").Value = "南區"
    ws.Range("B4").Value = 32
    ws.Range("C4").Value = 235
    ws.Range("D4").Value = 110

    ws.Range("A5").Value = "東區"
    ws.Range("B5").Value = 18
    ws.Range("C5").Value = 130
    ws.Range("D5").Value = 65

    ws.Range("A6").Value = "海外A"
    ws.Range("B6").Value = 45
    ws.Range("C6").Value = 340
    ws.Range("D6").Value = 150

    ws.Range("A7").Value = "海外B"
    ws.Range("B7").Value = 22
    ws.Range("C7").Value = 170
    ws.Range("D7").Value = 80

    ws.Range("A8").Value = "電商"
    ws.Range("B8").Value = 50
    ws.Range("C8").Value = 390
    ws.Range("D8").Value = 180

    ws.Columns("A:D").AutoFit
End Sub