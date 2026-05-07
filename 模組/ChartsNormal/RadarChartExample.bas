Option Explicit
Attribute VB_Name = "RadarChartExample"
'*************************************************************************************
'模組名稱: 雷達圖範例
'功能描述: 在 Excel 中建立雷達圖的範例程式
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
Sub TestRadarChart()
    Call CreateRadarChart("雷達圖範例")
End Sub

' 建立雷達圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateRadarChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    Call FillRadarChartData(ws)
    Set dataRange = ws.Range("A1:C7")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=440, _
        Height:=340)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlRadarMarkers

    cht.HasTitle = True
    cht.ChartTitle.Text = "部門能力評估雷達圖"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom
    cht.ChartStyle = 27

    MsgBox "雷達圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立雷達圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
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

' 輸入雷達圖範例資料
Private Sub FillRadarChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "指標"
    ws.Range("B1").Value = "業務部"
    ws.Range("C1").Value = "客服部"

    ws.Range("A2").Value = "效率"
    ws.Range("B2").Value = 85
    ws.Range("C2").Value = 78

    ws.Range("A3").Value = "品質"
    ws.Range("B3").Value = 80
    ws.Range("C3").Value = 92

    ws.Range("A4").Value = "成本"
    ws.Range("B4").Value = 76
    ws.Range("C4").Value = 83

    ws.Range("A5").Value = "創新"
    ws.Range("B5").Value = 88
    ws.Range("C5").Value = 70

    ws.Range("A6").Value = "協作"
    ws.Range("B6").Value = 82
    ws.Range("C6").Value = 90

    ws.Range("A7").Value = "穩定"
    ws.Range("B7").Value = 79
    ws.Range("C7").Value = 86

    ws.Columns("A:C").AutoFit
End Sub