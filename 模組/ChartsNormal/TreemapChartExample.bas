Attribute VB_Name = "TreemapChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: TreemapChartExample
'功能說明: 在 Excel 中建立樹狀圖（Treemap Chart）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestTreemapChart()
    Call CreateTreemapChart("樹狀圖範例")
End Sub

' 建立樹狀圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateTreemapChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    Call FillTreemapData(ws)
    Set dataRange = ws.Range("A1:C9")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=480, _
        Height:=320)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = 117 ' xlTreemap

    cht.HasTitle = True
    cht.ChartTitle.Text = "各部門費用佔比樹狀圖"

    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    MsgBox "樹狀圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立樹狀圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
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

' 填入樹狀圖範例資料
Private Sub FillTreemapData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "類別"
    ws.Range("B1").Value = "子類別"
    ws.Range("C1").Value = "金額"

    ws.Range("A2").Value = "行政部門"
    ws.Range("B2").Value = "辦公用品"
    ws.Range("C2").Value = 45000

    ws.Range("A3").Value = "行政部門"
    ws.Range("B3").Value = "差旅費"
    ws.Range("C3").Value = 32000

    ws.Range("A4").Value = "業務部門"
    ws.Range("B4").Value = "行銷廣告"
    ws.Range("C4").Value = 120000

    ws.Range("A5").Value = "業務部門"
    ws.Range("B5").Value = "客戶招待"
    ws.Range("C5").Value = 58000

    ws.Range("A6").Value = "研發部門"
    ws.Range("B6").Value = "設備採購"
    ws.Range("C6").Value = 98000

    ws.Range("A7").Value = "研發部門"
    ws.Range("B7").Value = "軟體授權"
    ws.Range("C7").Value = 67000

    ws.Range("A8").Value = "人資部門"
    ws.Range("B8").Value = "教育訓練"
    ws.Range("C8").Value = 40000

    ws.Range("A9").Value = "人資部門"
    ws.Range("B9").Value = "員工福利"
    ws.Range("C9").Value = 55000

    ws.Columns("A:C").AutoFit
End Sub
