Option Explicit
Attribute VB_Name = "StockChartExample"
'*************************************************************************************
'模組名稱: 股價圖範例
'功能描述: 在 Excel 中建立開高低收股價圖的範例程式
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
Sub TestStockChart()
    Call CreateStockChart("股價圖範例")
End Sub

' 建立股價圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateStockChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    Call FillStockChartData(ws)
    Set dataRange = ws.Range("A1:E11")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("G1").Left, _
        Top:=ws.Range("G1").Top, _
        Width:=520, _
        Height:=340)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlStockOHLC

    cht.HasTitle = True
    cht.ChartTitle.Text = "開高低收股價走勢"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "日期"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "價格"
    End With

    cht.ChartStyle = 2

    MsgBox "股價圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立股價圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
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

' 輸入股價圖範例資料
Private Sub FillStockChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "開盤"
    ws.Range("C1").Value = "最高"
    ws.Range("D1").Value = "最低"
    ws.Range("E1").Value = "收盤"

    ws.Range("A2").Value = DateSerial(2026, 4, 1)
    ws.Range("B2").Value = 101
    ws.Range("C2").Value = 108
    ws.Range("D2").Value = 99
    ws.Range("E2").Value = 106

    ws.Range("A3").Value = DateSerial(2026, 4, 2)
    ws.Range("B3").Value = 106
    ws.Range("C3").Value = 110
    ws.Range("D3").Value = 103
    ws.Range("E3").Value = 104

    ws.Range("A4").Value = DateSerial(2026, 4, 3)
    ws.Range("B4").Value = 104
    ws.Range("C4").Value = 112
    ws.Range("D4").Value = 102
    ws.Range("E4").Value = 111

    ws.Range("A5").Value = DateSerial(2026, 4, 6)
    ws.Range("B5").Value = 111
    ws.Range("C5").Value = 116
    ws.Range("D5").Value = 109
    ws.Range("E5").Value = 114

    ws.Range("A6").Value = DateSerial(2026, 4, 7)
    ws.Range("B6").Value = 114
    ws.Range("C6").Value = 118
    ws.Range("D6").Value = 112
    ws.Range("E6").Value = 117

    ws.Range("A7").Value = DateSerial(2026, 4, 8)
    ws.Range("B7").Value = 117
    ws.Range("C7").Value = 120
    ws.Range("D7").Value = 113
    ws.Range("E7").Value = 115

    ws.Range("A8").Value = DateSerial(2026, 4, 9)
    ws.Range("B8").Value = 115
    ws.Range("C8").Value = 121
    ws.Range("D8").Value = 114
    ws.Range("E8").Value = 119

    ws.Range("A9").Value = DateSerial(2026, 4, 10)
    ws.Range("B9").Value = 119
    ws.Range("C9").Value = 124
    ws.Range("D9").Value = 118
    ws.Range("E9").Value = 123

    ws.Range("A10").Value = DateSerial(2026, 4, 13)
    ws.Range("B10").Value = 123
    ws.Range("C10").Value = 126
    ws.Range("D10").Value = 120
    ws.Range("E10").Value = 121

    ws.Range("A11").Value = DateSerial(2026, 4, 14)
    ws.Range("B11").Value = 121
    ws.Range("C11").Value = 128
    ws.Range("D11").Value = 119
    ws.Range("E11").Value = 126

    ws.Range("A2:A11").NumberFormat = "yyyy/mm/dd"
    ws.Columns("A:E").AutoFit
End Sub