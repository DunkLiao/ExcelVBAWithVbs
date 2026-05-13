Attribute VB_Name = "ClearTextWrapFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearTextWrapFormatting
'功能說明: 清除選取範圍或整個工作表的文字換行格式設定
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

' 清除目前選取範圍的文字換行格式
Sub ClearTextWrapInSelection()
    Dim rng As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除文字換行格式的儲存格範圍。", vbExclamation, "提示"
        Exit Sub
    End If

    Set rng = Selection
    rng.WrapText = False
    rng.Rows.AutoFit

    MsgBox "已清除選取範圍的文字換行格式，共 " & rng.Cells.Count & " 個儲存格。", _
        vbInformation, "完成"
End Sub

' 清除整個工作表的文字換行格式
Sub ClearTextWrapInActiveSheet()
    Dim ws  As Worksheet
    Dim rng As Range

    Set ws = ActiveSheet

    Dim answer As Integer
    answer = MsgBox("確定要清除「" & ws.Name & "」工作表所有儲存格的文字換行格式嗎？", _
        vbYesNo + vbQuestion, "確認")

    If answer = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Set rng = ws.UsedRange
    rng.WrapText = False
    rng.Rows.AutoFit
    Application.ScreenUpdating = True

    MsgBox "已清除「" & ws.Name & "」工作表所有已使用範圍的文字換行格式。", _
        vbInformation, "完成"
End Sub

' 清除活頁簿所有工作表的文字換行格式
Sub ClearTextWrapInAllSheets()
    Dim wb  As Workbook
    Dim ws  As Worksheet

    Set wb = ThisWorkbook

    Dim answer As Integer
    answer = MsgBox("確定要清除活頁簿所有工作表的文字換行格式嗎？", _
        vbYesNo + vbQuestion, "確認")

    If answer = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.UsedRange.WrapText = False
        ws.UsedRange.Rows.AutoFit
        On Error GoTo 0
    Next ws
    Application.ScreenUpdating = True

    MsgBox "已清除所有工作表的文字換行格式。", vbInformation, "完成"
End Sub
