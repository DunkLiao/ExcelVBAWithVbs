Attribute VB_Name = "ClearHyperlinkFormatting"
Option Explicit

'*************************************************************************************
'模組名稱: ClearHyperlinkFormatting
'功能說明: 清除儲存格中的超連結及其藍色底線字型格式
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 清除選取範圍的超連結與其格式
Sub ClearHyperlinkFormatting()
    Dim targetRange As Range
    Dim hl          As Hyperlink
    Dim intCount    As Integer

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除超連結的儲存格範圍。", vbExclamation, "警告"
        Exit Sub
    End If

    Set targetRange = Selection
    intCount = targetRange.Hyperlinks.Count

    If intCount = 0 Then
        MsgBox "選取範圍內沒有超連結。", vbInformation, "提示"
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ' 刪除所有超連結
    targetRange.Hyperlinks.Delete

    ' 重設字型（移除藍色底線）
    With targetRange.Font
        .Color     = RGB(0, 0, 0)
        .Underline = xlUnderlineStyleNone
    End With

    MsgBox "已清除 " & intCount & " 個超連結及其字型格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除超連結時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除整張工作表所有超連結
Sub ClearAllHyperlinksInSheet()
    Dim ws       As Worksheet
    Dim intCount As Integer

    Set ws = ActiveSheet
    intCount = ws.Hyperlinks.Count

    If intCount = 0 Then
        MsgBox "工作表「" & ws.Name & "」內沒有超連結。", vbInformation, "提示"
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ws.Hyperlinks.Delete

    MsgBox "已清除工作表「" & ws.Name & "」的 " & intCount & " 個超連結。", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除工作表超連結時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub