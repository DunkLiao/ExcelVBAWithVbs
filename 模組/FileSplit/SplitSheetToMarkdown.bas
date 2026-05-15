Attribute VB_Name = "SplitSheetToMarkdown"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetToMarkdown
'功能說明: 將作用中工作表使用範圍匯出為 Markdown 表格檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunSplitSheetToMarkdown()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim filePath As Variant
    Dim markdownText As String
    Dim fileNumber As Integer

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
        MsgBox "目前工作表沒有可匯出的資料。", vbInformation, "提示"
        Exit Sub
    End If

    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=ws.Name & ".md", _
        FileFilter:="Markdown Files (*.md), *.md")

    If VarType(filePath) = vbBoolean Then
        If filePath = False Then Exit Sub
    End If

    markdownText = ConvertRangeToMarkdown(ws.UsedRange)
    fileNumber = FreeFile
    Open CStr(filePath) For Output As #fileNumber
    Print #fileNumber, markdownText;
    Close #fileNumber

    MsgBox "Markdown 檔案已匯出: " & CStr(filePath), vbInformation, "完成"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    On Error GoTo 0
    MsgBox "匯出 Markdown 時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function ConvertRangeToMarkdown(ByVal targetRange As Range) As String
    Dim rowCount As Long
    Dim colCount As Long
    Dim r As Long
    Dim c As Long
    Dim outputText As String

    rowCount = targetRange.Rows.Count
    colCount = targetRange.Columns.Count

    For c = 1 To colCount
        outputText = outputText & "| " & EscapeMarkdownText(CStr(targetRange.Cells(1, c).Text)) & " "
    Next c
    outputText = outputText & "|" & vbCrLf

    For c = 1 To colCount
        outputText = outputText & "| --- "
    Next c
    outputText = outputText & "|" & vbCrLf

    For r = 2 To rowCount
        For c = 1 To colCount
            outputText = outputText & "| " & EscapeMarkdownText(CStr(targetRange.Cells(r, c).Text)) & " "
        Next c
        outputText = outputText & "|" & vbCrLf
    Next r

    ConvertRangeToMarkdown = outputText
End Function

Private Function EscapeMarkdownText(ByVal sourceText As String) As String
    sourceText = Replace(sourceText, "|", "\|")
    sourceText = Replace(sourceText, vbCrLf, " ")
    sourceText = Replace(sourceText, vbCr, " ")
    sourceText = Replace(sourceText, vbLf, " ")
    EscapeMarkdownText = sourceText
End Function
