Attribute VB_Name = "MergeWithPivotSummary"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithPivotSummary
'功能說明: 將多個工作表資料合併後，自動建立樞紐分析表摘要的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點
Sub TestMergeWithPivotSummary()
    Call SetupPivotSampleData
    Call MergeAndBuildPivot("彙整資料", "樞紐摘要")
End Sub

' 合併所有來源工作表資料並建立樞紐分析表摘要
' mergeSheetName: 合併後的工作表名稱
' pivotSheetName: 樞紐分析表所在工作表名稱
Sub MergeAndBuildPivot(ByVal mergeSheetName As String, ByVal pivotSheetName As String)
    On Error GoTo ErrorHandler

    Dim wsMerge      As Worksheet
    Dim wsPivot      As Worksheet
    Dim ws           As Worksheet
    Dim destRow      As Long
    Dim srcLast      As Long
    Dim hasHeader    As Boolean
    Dim pc           As PivotCache
    Dim pt           As PivotTable
    Dim srcRange     As Range
    Dim mergeLastRow As Long
    Dim mergeLastCol As Long

    Set wsMerge = GetOrCreatePivotSummarySheet(mergeSheetName)
    wsMerge.Cells.Clear
    destRow = 1
    hasHeader = False

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> mergeSheetName And ws.Name <> pivotSheetName Then
            srcLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If srcLast >= 2 Then
                If Not hasHeader Then
                    ws.Rows(1).Copy Destination:=wsMerge.Rows(destRow)
                    destRow = destRow + 1
                    hasHeader = True
                End If
                ws.Range(ws.Cells(2, 1), ws.Cells(srcLast, _
                    ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)).Copy _
                    Destination:=wsMerge.Cells(destRow, 1)
                destRow = destRow + srcLast - 1
            End If
        End If
    Next ws

    wsMerge.Columns.AutoFit

    Set wsPivot = GetOrCreatePivotSummarySheet(pivotSheetName)
    wsPivot.Cells.Clear

    mergeLastRow = wsMerge.Cells(wsMerge.Rows.Count, 1).End(xlUp).Row
    mergeLastCol = wsMerge.Cells(1, wsMerge.Columns.Count).End(xlToLeft).Column
    Set srcRange = wsMerge.Range(wsMerge.Cells(1, 1), wsMerge.Cells(mergeLastRow, mergeLastCol))

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="MergePivot")

    With pt
        .PivotFields(1).Orientation = xlRowField
        .PivotFields(1).Position = 1
        If mergeLastCol >= 2 Then
            With .PivotFields(mergeLastCol)
                .Orientation = xlDataField
                .Function = xlSum
                .NumberFormat = "#,##0"
            End With
        End If
        .RowAxisLayout xlOutlineRow
    End With

    wsPivot.Columns.AutoFit

    Application.ScreenUpdating = True
    MsgBox "合併及樞紐摘要已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "合併建立樞紐時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 建立範例資料工作表（供測試用）
Private Sub SetupPivotSampleData()
    Dim ws As Worksheet
    Dim i  As Integer

    For i = 1 To 2
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("Q" & i)
        On Error GoTo 0
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = "Q" & i
        End If
        ws.Cells.Clear
        ws.Range("A1:C1").Value = Array("部門", "月份", "銷售額")
        ws.Range("A2:C2").Value = Array("業務部", i & "月", 120000 + i * 30000)
        ws.Range("A3:C3").Value = Array("行銀部", i & "月", 85000 + i * 15000)
        ws.Range("A4:C4").Value = Array("技術部", i & "月", 60000 + i * 10000)
    Next i
End Sub

' 取得或建立工作表
Private Function GetOrCreatePivotSummarySheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreatePivotSummarySheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreatePivotSummarySheet Is Nothing Then
        Set GetOrCreatePivotSummarySheet = ThisWorkbook.Worksheets.Add
        GetOrCreatePivotSummarySheet.Name = sheetName
    End If
End Function
