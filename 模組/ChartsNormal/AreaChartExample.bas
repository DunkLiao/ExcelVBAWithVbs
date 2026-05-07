Option Explicit
Attribute VB_Name = "AreaChartExample"
'*************************************************************************************
'模組名稱: 區域圖範例
'功能描述: 在 Excel 中建立區域圖的範例程式
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
Sub TestAreaChart()
    Call CreateAreaChart("區域圖範例")
End Sub

' 建立區域圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateAreaChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    Call FillAreaChartData(ws)
    Set dataRange = ws.Range("A1:D7")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("F1").Left, _
        Top:=ws.Range("F1").Top, _
        Width:=480, _
        Height:=320)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlAreaStacked

    cht.HasTitle = True
    cht.ChartTitle.Text = "各產品每月累積營收"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "月份"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "營收（萬元）"
    End With

    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom
    cht.ChartStyle = 13

    MsgBox "區域圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立區域圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
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

' 輸入區域圖範例資料
Private Sub FillAreaChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "產品A"
    ws.Range("C1").Value = "產品B"
    ws.Range("D1").Value = "產品C"

    ws.Range("A2").Value = "一月"
    ws.Range("B2").Value = 120
    ws.Range("C2").Value = 90
    ws.Range("D2").Value = 60

    ws.Range("A3").Value = "二月"
    ws.Range("B3").Value = 135
    ws.Range("C3").Value = 98
    ws.Range("D3").Value = 72

    ws.Range("A4").Value = "三月"
    ws.Range("B4").Value = 150
    ws.Range("C4").Value = 110
    ws.Range("D4").Value = 80

    ws.Range("A5").Value = "四月"
    ws.Range("B5").Value = 170
    ws.Range("C5").Value = 128
    ws.Range("D5").Value = 96

    ws.Range("A6").Value = "五月"
    ws.Range("B6").Value = 188
    ws.Range("C6").Value = 142
    ws.Range("D6").Value = 105

    ws.Range("A7").Value = "六月"
    ws.Range("B7").Value = 205
    ws.Range("C7").Value = 160
    ws.Range("D7").Value = 118

    ws.Columns("A:D").AutoFit
End Sub