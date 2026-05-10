Attribute VB_Name = "PivotWaterfallChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotWaterfallChartExample
'功能說明: 以樞紐分析資料為基礎，建立瀑布圖（Waterfall Chart）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestPivotWaterfallChart()
    Call CreatePivotWaterfallChart("樞紐瀑布圖範例")
End Sub

' 建立樞紐瀑布圖
' sheetName: 要建立圖表的工作表名稱
Sub CreatePivotWaterfallChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    Call FillWaterfallData(ws)
    Set dataRange = ws.Range("A1:B8")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=500, _
        Height:=340)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = 119 ' xlWaterfall

    cht.HasTitle = True
    cht.ChartTitle.Text = "各季度損益瀑布圖"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "項目"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "金額（萬元）"
    End With

    cht.HasLegend = False

    MsgBox "樞紐瀑布圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立瀑布圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
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

' 填入瀑布圖範例資料（損益變化）
Private Sub FillWaterfallData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "項目"
    ws.Range("B1").Value = "金額（萬元）"
    ws.Range("A1:B1").Font.Bold = True

    ws.Range("A2").Value = "期初餘額"
    ws.Range("B2").Value = 500

    ws.Range("A3").Value = "Q1 營收"
    ws.Range("B3").Value = 320

    ws.Range("A4").Value = "Q1 費用"
    ws.Range("B4").Value = -180

    ws.Range("A5").Value = "Q2 營收"
    ws.Range("B5").Value = 410

    ws.Range("A6").Value = "Q2 費用"
    ws.Range("B6").Value = -220

    ws.Range("A7").Value = "業外損益"
    ws.Range("B7").Value = -45

    ws.Range("A8").Value = "期末餘額"
    ws.Range("B8").Value = 785

    ws.Columns("A:B").AutoFit

    ws.Range("A10").Value = "說明：期末餘額為各項目加總後的結果"
    ws.Range("A10").Font.Italic = True
    ws.Range("A10").Font.Color = RGB(128, 128, 128)
End Sub
