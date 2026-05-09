Attribute VB_Name = "ClearBorderFormattingOnly"
Option Explicit

'*************************************************************************************
'模組名稱: ClearBorderFormattingOnly
'功能說明: 只清除儲存格框線格式，保留字型/填色/數字格式不變
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 清除選取範圍的所有框線
Sub ClearBorderFormattingOnly()
    Dim targetRange As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除框線的儲存格範圍。", vbExclamation, "警告"
        Exit Sub
    End If

    Set targetRange = Selection

    On Error GoTo ErrHandler

    targetRange.Borders(xlEdgeLeft).LineStyle   = xlNone
    targetRange.Borders(xlEdgeRight).LineStyle  = xlNone
    targetRange.Borders(xlEdgeTop).LineStyle    = xlNone
    targetRange.Borders(xlEdgeBottom).LineStyle = xlNone
    targetRange.Borders(xlInsideHorizontal).LineStyle = xlNone
    targetRange.Borders(xlInsideVertical).LineStyle   = xlNone

    MsgBox "已清除選取範圍的框線格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除框線時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除整張工作表所有框線
Sub ClearAllBordersInSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    On Error GoTo ErrHandler

    ws.Cells.Borders(xlEdgeLeft).LineStyle         = xlNone
    ws.Cells.Borders(xlEdgeRight).LineStyle        = xlNone
    ws.Cells.Borders(xlEdgeTop).LineStyle          = xlNone
    ws.Cells.Borders(xlEdgeBottom).LineStyle       = xlNone
    ws.Cells.Borders(xlInsideHorizontal).LineStyle = xlNone
    ws.Cells.Borders(xlInsideVertical).LineStyle   = xlNone

    MsgBox "已清除工作表「" & ws.Name & "」的所有框線。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除工作表框線時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub