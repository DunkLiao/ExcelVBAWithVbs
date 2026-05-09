Attribute VB_Name = "ClearFillColorOnly"
Option Explicit

'*************************************************************************************
'模組名稱: ClearFillColorOnly
'功能說明: 只清除儲存格填滿色彩（背景色），保留其他格式不變
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 清除選取範圍的填滿色彩
Sub ClearFillColorOnly()
    Dim targetRange As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除填滿色彩的儲存格範圍。", vbExclamation, "警告"
        Exit Sub
    End If

    Set targetRange = Selection

    On Error GoTo ErrHandler

    targetRange.Interior.ColorIndex = xlNone

    MsgBox "已清除選取範圍的填滿色彩（共 " & targetRange.Cells.Count & " 個儲存格）。", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除填滿色彩時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除整張工作表的填滿色彩
Sub ClearAllFillColorsInSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    On Error GoTo ErrHandler

    ws.UsedRange.Interior.ColorIndex = xlNone

    MsgBox "已清除工作表「" & ws.Name & "」所有使用範圍的填滿色彩。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除填滿色彩時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除特定顏色的儲存格填滿（跳過其他顏色）
Sub ClearSpecificFillColor()
    Dim ws         As Worksheet
    Dim cell       As Range
    Dim targetColor As Long
    Dim intCleaned As Integer

    Set ws = ActiveSheet
    ' 以黃色 (RGB 255,255,0) 為目標
    targetColor = RGB(255, 255, 0)
    intCleaned = 0

    On Error GoTo ErrHandler

    For Each cell In ws.UsedRange
        If cell.Interior.Color = targetColor Then
            cell.Interior.ColorIndex = xlNone
            intCleaned = intCleaned + 1
        End If
    Next cell

    MsgBox "已清除 " & intCleaned & " 個黃色填滿儲存格的背景色。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除特定色彩時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub