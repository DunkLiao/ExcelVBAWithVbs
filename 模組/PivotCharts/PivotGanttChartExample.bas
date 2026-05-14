Attribute VB_Name = "PivotGanttChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotGanttChartExample
'功能說明: 以樞紐分析表為資料來源，建立甘特式堆疊橫條圖的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點
Sub TestPivotGanttChart()
    Call SetupGanttSourceData
    Call CreatePivotGanttChart("甘特圖來源", "甘特樞紐圖")
End Sub

' 建立樞紐甘特式堆疊橫條圖
' dataSheetName: 原始資料工作表名稱
' pivotSheetName: 樞紐圖輸出工作表名稱
Sub CreatePivotGanttChart(ByVal dataSheetName As String, ByVal pivotSheetName As String)
    On Error GoTo ErrorHandler

    Dim wsData   As Worksheet
    Dim wsPivot  As Worksheet
    Dim pc       As PivotCache
    Dim pt       As PivotTable
    Dim cht      As Chart
    Dim co       As ChartObject
    Dim lastRow  As Long
    Dim lastCol  As Long
    Dim srcRange As Range

    Application.ScreenUpdating = False

    Set wsData = ThisWorkbook.Worksheets(dataSheetName)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set srcRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcRange)

    Set wsPivot = GetOrCreateGanttSheet(pivotSheetName)
    wsPivot.Cells.Clear

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="GanttPivot")

    With pt
        .PivotFields("任務名稱").Orientation = xlRowField
        .PivotFields("任務名稱").Position = 1
        With .PivotFields("開始天")
            .Orientation = xlDataField
            .Function = xlMin
            .Caption = "開始天"
        End With
        With .PivotFields("持續天數")
            .Orientation = xlDataField
            .Function = xlSum
            .Caption = "持續天數"
        End With
        .ColumnGrand = False
        .RowGrand = False
    End With

    wsPivot.Columns.AutoFit

    Set co = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("E2").Left, _
        Top:=wsPivot.Range("E2").Top, _
        Width:=500, Height:=280)
    Set cht = co.Chart

    cht.SetSourceData Source:=pt.DataBodyRange
    cht.ChartType = xlBarStacked

    cht.HasTitle = True
    cht.ChartTitle.Text = "專案任務甘特圖（樞紐橫條圖）"

    ' 將第一個系列（開始天）設為透明（模擬甘特圖效果）
    With cht.SeriesCollection(1)
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With

    With cht.SeriesCollection(2)
        .Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
    End With

    Application.ScreenUpdating = True
    MsgBox "樞紐甘特圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "建立樞紐甘特圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 建立甘特圖範例資料（任務名稱、開始天、持續天數）
Private Sub SetupGanttSourceData()
    Dim ws As Worksheet
    Set ws = GetOrCreateGanttSheet("甘特圖來源")
    ws.Cells.Clear

    ws.Range("A1:C1").Value = Array("任務名稱", "開始天", "持續天數")
    ws.Range("A2:C2").Value = Array("需求分析", 0, 5)
    ws.Range("A3:C3").Value = Array("系統設計", 5, 7)
    ws.Range("A4:C4").Value = Array("程式開發", 12, 14)
    ws.Range("A5:C5").Value = Array("測試驗收", 26, 5)
    ws.Range("A6:C6").Value = Array("上線部署", 31, 3)
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateGanttSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateGanttSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateGanttSheet Is Nothing Then
        Set GetOrCreateGanttSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateGanttSheet.Name = sheetName
    End If
End Function
