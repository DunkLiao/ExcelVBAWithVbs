Attribute VB_Name = "ClearSelectedCellsFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearSelectedCellsFormatting
'功能說明: 清除目前選取儲存格的所有格式，保留儲存格內容
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub ClearSelectedCellsFormatting()
    Dim selRange As Range
    Dim confirm As Integer

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取儲存格範圍後再執行。", vbExclamation, "提示"
        Exit Sub
    End If

    Set selRange = Selection

    confirm = MsgBox("確定要清除選取範圍的所有格式嗎？" & vbCrLf & _
        "（內容將保留，格式將全部清除）", vbYesNo + vbQuestion, "確認")

    If confirm = vbNo Then Exit Sub

    selRange.ClearFormats

    MsgBox "已清除選取範圍的所有格式。", vbInformation, "完成"
End Sub

Sub ClearSelectedCellsFormattingKeepColor()
    Dim selRange As Range
    Dim cell As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取儲存格範圍後再執行。", vbExclamation, "提示"
        Exit Sub
    End If

    Set selRange = Selection

    For Each cell In selRange
        With cell.Font
            .Bold = False
            .Italic = False
            .Underline = xlUnderlineStyleNone
            .Color = RGB(0, 0, 0)
            .Size = 11
            .Name = "新細明體"
        End With
        cell.Borders.LineStyle = xlNone
        cell.NumberFormat = "General"
    Next cell

    MsgBox "已清除字型與框線格式（填色保留）。", vbInformation, "完成"
End Sub
