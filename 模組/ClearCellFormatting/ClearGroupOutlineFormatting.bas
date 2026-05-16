Attribute VB_Name = "ClearGroupOutlineFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearGroupOutlineFormatting
'功能說明: 清除工作表的列/欄群組（大綱）設定，並可選擇性地
'          同時清除自動小計、展開/折疊狀態或所有工作表的群組
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

' 清除作用中工作表的所有列/欄群組
Sub ClearAllGroupsInActiveSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 移除所有列群組
    ws.Rows.ClearOutline

    ' 移除所有欄群組（Rows.ClearOutline 已一併清除，此處另行確認）
    ws.Columns.ClearOutline

    ' 清除大綱設定
    ws.Outline.SummaryRow = xlSummaryAbove
    ws.Outline.SummaryColumn = xlSummaryOnLeft

    MsgBox "已清除「" & ws.Name & "」的所有群組（大綱）設定！", vbInformation, "完成"
End Sub

' 清除活頁簿中所有工作表的群組
Sub ClearAllGroupsInAllSheets()
    Dim ws As Worksheet
    Dim count As Long
    count = 0

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Rows.ClearOutline
        ws.Columns.ClearOutline
        On Error GoTo 0
        count = count + 1
    Next ws

    MsgBox "已清除全部 " & count & " 張工作表的群組（大綱）設定！", vbInformation, "完成"
End Sub

' 移除自動小計並清除群組（SubTotal 常伴隨群組產生）
Sub RemoveSubtotalAndClearGroups()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "工作表沒有資料。", vbExclamation
        Exit Sub
    End If

    ' 移除自動小計
    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).RemoveSubtotal
    On Error GoTo 0

    ' 清除群組大綱
    ws.Rows.ClearOutline
    ws.Columns.ClearOutline

    MsgBox "已移除自動小計並清除群組設定！", vbInformation, "完成"
End Sub

' 示範：先建立群組，再清除
Sub DemoCreateAndClearGroups()
    Dim ws As Worksheet
    Set ws = GetOrCreateGroupSheet("群組示範")
    ws.Cells.Clear

    ' 填入範例資料
    ws.Range("A1:C1").Value = Array("月份", "業務員", "銷售量")
    Dim i As Integer
    For i = 1 To 9
        ws.Cells(i + 1, 1).Value = "第" & ((i - 1) \ 3 + 1) & "季"
        ws.Cells(i + 1, 2).Value = "業務" & i
        ws.Cells(i + 1, 3).Value = (i * 15000)
    Next i

    ' 建立列群組（示範用）
    ws.Rows("2:4").Group
    ws.Rows("5:7").Group
    ws.Rows("8:10").Group

    MsgBox "已建立三個列群組（第2~4、5~7、8~10列）。" & vbCrLf & _
           "按確定後將清除所有群組。", vbInformation, "步驟1"

    ' 清除群組
    ws.Rows.ClearOutline
    ws.Columns.ClearOutline

    MsgBox "所有群組已清除！", vbInformation, "完成"
End Sub

Private Function GetOrCreateGroupSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateGroupSheet = ws
End Function
