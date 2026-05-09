Attribute VB_Name = "ClearRowColumnFormatting"
Option Explicit

'*************************************************************************************
'模組名稱: ClearRowColumnFormatting
'功能說明: 清除整列或整欄的格式設定（字型/填色/框線/數字格式）
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 清除目前所在列的格式
Sub ClearCurrentRowFormatting()
    Dim ws  As Worksheet
    Dim lngRow As Long

    Set ws = ActiveSheet
    lngRow = ActiveCell.Row

    On Error GoTo ErrHandler

    ws.Rows(lngRow).ClearFormats
    ws.Rows(lngRow).FormatConditions.Delete

    MsgBox "已清除第 " & lngRow & " 列的所有格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除列格式時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除目前所在欄的格式
Sub ClearCurrentColumnFormatting()
    Dim ws     As Worksheet
    Dim lngCol As Long

    Set ws = ActiveSheet
    lngCol = ActiveCell.Column

    On Error GoTo ErrHandler

    ws.Columns(lngCol).ClearFormats
    ws.Columns(lngCol).FormatConditions.Delete

    MsgBox "已清除第 " & lngCol & " 欄的所有格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除欄格式時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 依使用者輸入列號範圍批次清除格式
Sub ClearRowRangeFormatting()
    Dim ws         As Worksheet
    Dim strInput   As String
    Dim lngStart   As Long
    Dim lngEnd     As Long

    Set ws = ActiveSheet

    strInput = InputBox("請輸入要清除格式的列號範圍（例如：3:8）：", "輸入列範圍")
    If strInput = "" Then
        MsgBox "已取消操作。", vbInformation, "提示"
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ws.Rows(strInput).ClearFormats
    ws.Rows(strInput).FormatConditions.Delete

    MsgBox "已清除第 " & strInput & " 列的格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除列範圍格式時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub