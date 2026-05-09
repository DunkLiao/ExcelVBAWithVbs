Attribute VB_Name = "ClearCommentAndFormatting"
Option Explicit

'*************************************************************************************
'模組名稱: ClearCommentAndFormatting
'功能說明: 同時清除儲存格的註解（批註）與格式設定
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 清除選取範圍的註解與格式
Sub ClearCommentAndFormatting()
    Dim targetRange As Range
    Dim intComments As Integer

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除的儲存格範圍。", vbExclamation, "警告"
        Exit Sub
    End If

    Set targetRange = Selection
    intComments = targetRange.SpecialCells(xlCellTypeComments).Count

    On Error GoTo ErrHandler

    targetRange.ClearFormats
    targetRange.ClearComments

    MsgBox "已清除選取範圍的格式與 " & intComments & " 個註解。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    ' 若無註解仍繼續清除格式
    targetRange.ClearFormats
    MsgBox "已清除選取範圍的格式（無找到可清除的註解）。", vbInformation, "完成"
End Sub

' 清除整張工作表所有註解與使用範圍格式
Sub ClearAllCommentsAndFormatsInSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    ws.UsedRange.ClearFormats
    ws.UsedRange.ClearComments

    Application.ScreenUpdating = True
    MsgBox "已清除工作表「" & ws.Name & "」使用範圍的所有格式與註解。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "清除時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub