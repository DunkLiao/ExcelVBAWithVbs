Attribute VB_Name = "GanttChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: GanttChartExample
'功能說明: 以堆疊橫條圖模擬甘特圖，展示專案排程視覺化範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestGanttChart()
    Call CreateGanttChart
End Sub

' 建立甘特圖範例
Sub CreateGanttChart()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range
    Dim sheetName As String

    sheetName = "甘特圖範例"

    ' 取得或建立工作表
    Set ws = GetOrCreateSheet(ThisWorkbook, sheetName)

    ' 填入標題列
    ws.Range("A1").Value = "任務名稱"
    ws.Range("B1").Value = "開始天數"
    ws.Range("C1").Value = "持續天數"
    ws.Range("A1:C1").Font.Bold = True

    ' 填入任務資料
    ws.Range("A2").Value = "需求分析"
    ws.Range("B2").Value = 0
    ws.Range("C2").Value = 5

    ws.Range("A3").Value = "系統設計"
    ws.Range("B3").Value = 5
    ws.Range("C3").Value = 7

    ws.Range("A4").Value = "程式開發"
    ws.Range("B4").Value = 12
    ws.Range("C4").Value = 10

    ws.Range("A5").Value = "測試驗收"
    ws.Range("B5").Value = 22
    ws.Range("C5").Value = 5

    ws.Range("A6").Value = "上線部署"
    ws.Range("B6").Value = 27
    ws.Range("C6").Value = 3

    ws.Columns("A:C").AutoFit

    ' 設定資料範圍
    Set dataRange = ws.Range("A1:C6")

    ' 新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=480, _
        Height:=300)

    Set cht = chartObj.Chart

    ' 設定來源資料
    cht.SetSourceData Source:=dataRange

    ' 使用堆疊橫條圖
    cht.ChartType = xlBarStacked

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "專案甘特圖"

    ' 將第一數列（開始天數）設為透明，模擬甘特圖效果
    With cht.SeriesCollection(1).Format.Fill
        .Visible = msoFalse
    End With

    ' 設定圖例
    cht.HasLegend = True

    MsgBox "甘特圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立甘特圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表並清空內容
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
