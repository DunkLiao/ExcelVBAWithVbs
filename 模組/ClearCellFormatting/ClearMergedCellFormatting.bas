Attribute VB_Name = "ClearMergedCellFormatting"
Option Explicit

'*************************************************************************************
'模組名稱: ClearMergedCellFormatting
'功能說明: 取消合併儲存格並清除其格式設定
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 取消選取範圍的合併並清除格式
Sub UnmergeAndClearFormatting()
    Dim targetRange As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要取消合併並清除格式的儲存格範圍。", vbExclamation, "警告"
        Exit Sub
    End If

    Set targetRange = Selection

    On Error GoTo ErrHandler

    targetRange.UnMerge
    targetRange.ClearFormats

    MsgBox "已取消合併並清除格式（範圍：" & targetRange.Address & "）。", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "取消合併或清除格式時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 掃描整張工作表，找出所有合併儲存格並取消合併後清除格式
Sub ClearAllMergedCellFormatsInSheet()
    Dim ws        As Worksheet
    Dim cell      As Range
    Dim intCount  As Integer
    Dim mergedArea As Range

    Set ws = ActiveSheet
    intCount = 0

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    For Each cell In ws.UsedRange
        If cell.MergeCells Then
            If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                Set mergedArea = cell.MergeArea
                mergedArea.UnMerge
                mergedArea.ClearFormats
                intCount = intCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "已取消並清除 " & intCount & " 個合併儲存格區域的格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "處理合併儲存格時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub