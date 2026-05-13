Attribute VB_Name = "ClearIndentFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearIndentFormatting
'功能說明: 清除選取範圍或整個工作表中所有儲存格的縮排與對齊格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub ClearIndentInSelection()
    On Error GoTo ErrHandler
    Dim rng  As Range
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除縮排的儲存格範圍。", vbExclamation, "提示"
        Exit Sub
    End If
    Set rng = Selection
    For Each cell In rng.Cells
        cell.IndentLevel = 0
        cell.HorizontalAlignment = xlGeneral
        cell.VerticalAlignment   = xlBottom
    Next cell
    MsgBox "已清除選取範圍的縮排格式，共 " & rng.Cells.Count & " 個儲存格。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Public Sub ClearIndentInActiveSheet()
    On Error GoTo ErrHandler
    Dim ws   As Worksheet
    Dim rng  As Range
    Dim cell As Range
    Set ws  = ActiveSheet
    Set rng = ws.UsedRange
    For Each cell In rng.Cells
        cell.IndentLevel = 0
        cell.HorizontalAlignment = xlGeneral
    Next cell
    MsgBox "已清除工作表「" & ws.Name & "」所有儲存格的縮排格式。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Public Sub CreateIndentSampleData()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("縮排格式範例")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "縮排格式範例"
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "大類" : ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "中類" : ws.Range("A2").IndentLevel = 1
    ws.Range("A3").Value = "小類" : ws.Range("A3").IndentLevel = 2
    ws.Range("A4").Value = "細項" : ws.Range("A4").IndentLevel = 3
    ws.Columns("A").AutoFit
    ws.Activate
    MsgBox "已建立含縮排格式的範例資料。請執行 ClearIndentInActiveSheet 清除縮排。", vbInformation, "完成"
End Sub

