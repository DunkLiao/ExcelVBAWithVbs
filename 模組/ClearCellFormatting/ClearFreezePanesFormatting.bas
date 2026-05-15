Attribute VB_Name = "ClearFreezePanesFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearFreezePanesFormatting
'功能說明: 清除所有工作表的凍結窗格與分割設定並回報處理數量
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunClearFreezePanesFormatting()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim originalSheet As Worksheet
    Dim originalVisibility As XlSheetVisibility
    Dim processedCount As Long

    If ThisWorkbook.Windows.Count = 0 Then Exit Sub

    Set originalSheet = ActiveSheet
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        originalVisibility = ws.Visible
        If originalVisibility <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If

        ws.Activate
        If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False
        If ActiveWindow.Split Then ActiveWindow.Split = False
        processedCount = processedCount + 1

        If originalVisibility <> xlSheetVisible Then
            ws.Visible = originalVisibility
        End If
    Next ws

    originalSheet.Activate
    Application.ScreenUpdating = True

    MsgBox "已處理 " & processedCount & " 張工作表的凍結窗格與分割設定。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not originalSheet Is Nothing Then originalSheet.Activate
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "清除凍結窗格與分割設定時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub
