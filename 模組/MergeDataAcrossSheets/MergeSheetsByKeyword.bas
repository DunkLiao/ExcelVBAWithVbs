Option Explicit

' 合併工作表名稱包含指定關鍵字的資料，適合彙整分店或月份工作表。
Public Sub MergeSheetsByKeywordExample()
    On Error GoTo ErrHandler

    Dim keyword As String
    Dim targetWs As Worksheet

    keyword = CStr(Application.InputBox("請輸入工作表名稱關鍵字", "跨表合併", "分店", Type:=2))
    If Len(keyword) = 0 Then Exit Sub

    Set targetWs = GetOrCreateMergeWorksheet("關鍵字合併結果")
    Call MergeSheetsByKeyword(keyword, targetWs)

    MsgBox "跨表合併完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "跨表合併失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub MergeSheetsByKeyword(ByVal keyword As String, ByVal targetWs As Worksheet)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim targetRow As Long
    Dim copyStartRow As Long

    targetWs.Cells.Clear
    targetRow = 1

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> targetWs.Name And InStr(1, ws.Name, keyword, vbTextCompare) > 0 Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lastRow > 0 And lastCol > 0 Then
                If targetRow = 1 Then
                    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy Destination:=targetWs.Cells(targetRow, 1)
                    targetWs.Cells(targetRow, lastCol + 1).Value = "來源工作表"
                    targetRow = targetRow + 1
                    copyStartRow = 2
                Else
                    copyStartRow = 2
                End If

                If lastRow >= copyStartRow Then
                    ws.Range(ws.Cells(copyStartRow, 1), ws.Cells(lastRow, lastCol)).Copy Destination:=targetWs.Cells(targetRow, 1)
                    targetWs.Range(targetWs.Cells(targetRow, lastCol + 1), targetWs.Cells(targetRow + lastRow - copyStartRow, lastCol + 1)).Value = ws.Name
                    targetRow = targetRow + lastRow - copyStartRow + 1
                End If
            End If
        End If
    Next ws

    targetWs.Columns.AutoFit
End Sub

Private Function GetOrCreateMergeWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateMergeWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateMergeWorksheet Is Nothing Then
        Set GetOrCreateMergeWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateMergeWorksheet.Name = sheetName
    End If
End Function