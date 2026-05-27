Option Explicit
Attribute VB_Name = "ClearCellCommentFormatting"
'*************************************************************************************
'模組名稱: 清除儲存格批注
'功能說明: 清除使用中工作表（或選取範圍）內所有儲存格的批注（Comment）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub ClearAllCommentsInSheet()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim commentCount As Long
    commentCount = ws.Comments.Count

    If commentCount = 0 Then
        MsgBox "此工作表沒有任何批注。", vbInformation, "提示"
        Exit Sub
    End If

    Dim result As VbMsgBoxResult
    result = MsgBox("確定要刪除此工作表中全部 " & commentCount & " 個批注嗎？", _
        vbQuestion + vbYesNo, "確認")

    If result = vbNo Then Exit Sub

    ws.Cells.ClearComments

    MsgBox "已清除 " & commentCount & " 個批注。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "清除批注時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Sub ClearCommentsInSelection()
    On Error GoTo ErrorHandler

    Dim sel As Range
    On Error Resume Next
    Set sel = Selection
    On Error GoTo ErrorHandler

    If sel Is Nothing Then
        MsgBox "請先選取範圍。", vbExclamation, "提示"
        Exit Sub
    End If

    Dim cnt As Long
    Dim c As Range
    For Each c In sel.Cells
        If Not c.Comment Is Nothing Then cnt = cnt + 1
    Next c

    If cnt = 0 Then
        MsgBox "選取範圍內沒有批注。", vbInformation, "提示"
        Exit Sub
    End If

    sel.ClearComments
    MsgBox "已清除選取範圍內 " & cnt & " 個批注。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "清除批注時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
