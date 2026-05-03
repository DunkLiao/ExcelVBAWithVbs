Option Explicit
Attribute VB_Name = "ComboChartExample"
'*************************************************************************************
'模組名稱: 組合圖範例
'功能描述: 在 Excel 中建立直條圖與折線圖組合圖的範例程式
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
Sub TestComboChart()
    Call CreateComboChart("組合圖範例")
End Sub

' 建立組合圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateComboChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    Call FillComboChartData(ws)
    Set dataRange = ws.Range("A1:C7")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=500, _
        Height:=340)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlColumnClustered

    With cht.SeriesCollection(2)
        .ChartType = xlLineMarkers
        .AxisGroup = xlSecondary
        .HasDataLabels = True
    End With

    cht.HasTitle = True
    cht.ChartTitle.Text = "營收與毛利率組合圖"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "月份"
    End With

    With cht.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "營收（萬元）"
    End With

    With cht.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = "毛利率"
        .TickLabels.NumberFormat = "0%"
    End With

    cht.ChartStyle = 5

    MsgBox "組合圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立組合圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
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

' 輸入組合圖範例資料
Private Sub FillComboChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "營收"
    ws.Range("C1").Value = "毛利率"

    ws.Range("A2").Value = "一月"
    ws.Range("B2").Value = 420
    ws.Range("C2").Value = 0.31

    ws.Range("A3").Value = "二月"
    ws.Range("B3").Value = 455
    ws.Range("C3").Value = 0.33

    ws.Range("A4").Value = "三月"
    ws.Range("B4").Value = 510
    ws.Range("C4").Value = 0.35

    ws.Range("A5").Value = "四月"
    ws.Range("B5").Value = 498
    ws.Range("C5").Value = 0.34

    ws.Range("A6").Value = "五月"
    ws.Range("B6").Value = 560
    ws.Range("C6").Value = 0.37

    ws.Range("A7").Value = "六月"
    ws.Range("B7").Value = 610
    ws.Range("C7").Value = 0.39

    ws.Range("C2:C7").NumberFormat = "0%"
    ws.Columns("A:C").AutoFit
End Sub