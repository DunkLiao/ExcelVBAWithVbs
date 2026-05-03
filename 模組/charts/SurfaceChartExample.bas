Option Explicit
Attribute VB_Name = "SurfaceChartExample"
'*************************************************************************************
'模組名稱: 曲面圖範例
'功能描述: 在 Excel 中建立三維曲面圖的範例程式
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
Sub TestSurfaceChart()
    Call CreateSurfaceChart("曲面圖範例")
End Sub

' 建立曲面圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateSurfaceChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    Call FillSurfaceChartData(ws)
    Set dataRange = ws.Range("A1:F6")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("H1").Left, _
        Top:=ws.Range("H1").Top, _
        Width:=520, _
        Height:=360)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlSurface

    cht.HasTitle = True
    cht.ChartTitle.Text = "價格與數量對利潤的影響"
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.ChartStyle = 14

    MsgBox "曲面圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立曲面圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
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

' 輸入曲面圖範例資料
Private Sub FillSurfaceChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "數量\價格"
    ws.Range("B1").Value = 80
    ws.Range("C1").Value = 90
    ws.Range("D1").Value = 100
    ws.Range("E1").Value = 110
    ws.Range("F1").Value = 120

    ws.Range("A2").Value = 100
    ws.Range("B2").Value = 120
    ws.Range("C2").Value = 180
    ws.Range("D2").Value = 230
    ws.Range("E2").Value = 260
    ws.Range("F2").Value = 280

    ws.Range("A3").Value = 200
    ws.Range("B3").Value = 260
    ws.Range("C3").Value = 340
    ws.Range("D3").Value = 420
    ws.Range("E3").Value = 470
    ws.Range("F3").Value = 500

    ws.Range("A4").Value = 300
    ws.Range("B4").Value = 330
    ws.Range("C4").Value = 455
    ws.Range("D4").Value = 570
    ws.Range("E4").Value = 640
    ws.Range("F4").Value = 690

    ws.Range("A5").Value = 400
    ws.Range("B5").Value = 360
    ws.Range("C5").Value = 540
    ws.Range("D5").Value = 700
    ws.Range("E5").Value = 820
    ws.Range("F5").Value = 900

    ws.Range("A6").Value = 500
    ws.Range("B6").Value = 340
    ws.Range("C6").Value = 590
    ws.Range("D6").Value = 790
    ws.Range("E6").Value = 960
    ws.Range("F6").Value = 1080

    ws.Columns("A:F").AutoFit
End Sub