Attribute VB_Name = "ClearColumnWidthFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearColumnWidthFormatting
'功能說明: 重設工作表中所有欄寬至 Excel 預設值（8.43 字元寬），清除自訂欄寬設定
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Private Const DEFAULT_COLUMN_WIDTH As Double = 8.43

Sub TestClearColumnWidth()
    Call ClearColumnWidthFormatting(ActiveSheet)
End Sub

Sub ClearColumnWidthFormatting(ByVal ws As Worksheet)
    Dim totalCols  As Integer
    Dim i          As Integer
    Dim resetCount As Integer

    Application.ScreenUpdating = False
    resetCount = 0
    totalCols = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
    If totalCols < 1 Then totalCols = 1

    For i = 1 To totalCols
        If Abs(ws.Columns(i).ColumnWidth - DEFAULT_COLUMN_WIDTH) > 0.01 Then
            ws.Columns(i).ColumnWidth = DEFAULT_COLUMN_WIDTH
            resetCount = resetCount + 1
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "欄寬重設完成！" & vbCrLf & _
           "共重設 " & resetCount & " 個欄位為預設寬度（" & DEFAULT_COLUMN_WIDTH & " 字元）。", _
           vbInformation, "完成"
End Sub

Sub ClearSelectedColumnWidth()
    Dim sel        As Range
    Dim col        As Range
    Dim resetCount As Integer

    Set sel = Selection
    If sel Is Nothing Then
        MsgBox "請先選取要重設欄寬的範圍。", vbExclamation, "提示"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    resetCount = 0
    For Each col In sel.EntireColumn.Columns
        If Abs(col.ColumnWidth - DEFAULT_COLUMN_WIDTH) > 0.01 Then
            col.ColumnWidth = DEFAULT_COLUMN_WIDTH
            resetCount = resetCount + 1
        End If
    Next col
    Application.ScreenUpdating = True
    MsgBox "選取範圍欄寬重設完成！共重設 " & resetCount & " 個欄位。", _
           vbInformation, "完成"
End Sub

Sub ClearAllSheetsColumnWidth()
    Dim ws         As Worksheet
    Dim totalReset As Long
    Dim wsCount    As Integer
    Dim i          As Integer
    Dim totalCols  As Integer

    Application.ScreenUpdating = False
    totalReset = 0
    wsCount = 0
    For Each ws In ThisWorkbook.Worksheets
        totalCols = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        If totalCols < 1 Then totalCols = 1
        For i = 1 To totalCols
            If Abs(ws.Columns(i).ColumnWidth - DEFAULT_COLUMN_WIDTH) > 0.01 Then
                ws.Columns(i).ColumnWidth = DEFAULT_COLUMN_WIDTH
                totalReset = totalReset + 1
            End If
        Next i
        wsCount = wsCount + 1
    Next ws
    Application.ScreenUpdating = True
    MsgBox "全活頁簿欄寬重設完成！" & vbCrLf & _
           "共處理 " & wsCount & " 張工作表，重設 " & totalReset & " 個欄位。", _
           vbInformation, "完成"
End Sub
