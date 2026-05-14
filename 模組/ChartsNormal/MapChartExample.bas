Attribute VB_Name = "MapChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: MapChartExample
'功能說明: 在 Excel 中建立地圖圖表（filled map）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點
Sub TestMapChart()
    Call CreateMapChart("地圖圖表範例")
End Sub

' 建立地圖圖表
' sheetName: 要建立圖表的工作表名稱
Sub CreateMapChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws        As Worksheet
    Dim chartObj  As ChartObject
    Dim cht       As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateMapSheet(sheetName)
    ws.Cells.Clear

    Call FillMapChartData(ws)
    Set dataRange = ws.Range("A1:B7")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=480, _
        Height:=320)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlFilledMap

    cht.HasTitle = True
    cht.ChartTitle.Text = "各縣市銷售金額分布"

    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    MsgBox "地圖圖表已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立地圖圖表時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateMapSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateMapSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateMapSheet Is Nothing Then
        Set GetOrCreateMapSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateMapSheet.Name = sheetName
    End If
End Function

' 填入地圖圖表範例資料（縣市與銷售額）
Private Sub FillMapChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "縣市"
    ws.Range("B1").Value = "銷售金額"

    ws.Range("A2").Value = "台北市"
    ws.Range("B2").Value = 520000

    ws.Range("A3").Value = "新北市"
    ws.Range("B3").Value = 430000

    ws.Range("A4").Value = "桃園市"
    ws.Range("B4").Value = 310000

    ws.Range("A5").Value = "台中市"
    ws.Range("B5").Value = 280000

    ws.Range("A6").Value = "台南市"
    ws.Range("B6").Value = 195000

    ws.Range("A7").Value = "高雄市"
    ws.Range("B7").Value = 240000

    ws.Range("B2:B7").NumberFormat = "#,##0"
    ws.Columns("A:B").AutoFit
End Sub
