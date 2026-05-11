Attribute VB_Name = "MergeWithReferenceFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithReferenceFormulas
'功能說明: 以儲存格參照公式方式合併多張工作表資料至彙整表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub MergeWithReferenceFormulas()
    Dim summaryWs As Worksheet
    Dim ws As Worksheet
    Dim destRow As Long
    Dim srcLastRow As Long
    Dim srcLastCol As Long
    Dim i As Long
    Dim c As Integer
    Dim isFirst As Boolean

    Const SUMMARY_SHEET As String = "彙整_參照"

    On Error Resume Next
    Set summaryWs = ThisWorkbook.Worksheets(SUMMARY_SHEET)
    On Error GoTo 0

    If summaryWs Is Nothing Then
        Set summaryWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        summaryWs.Name = SUMMARY_SHEET
    Else
        summaryWs.Cells.Clear
    End If

    destRow = 1
    isFirst = True

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SUMMARY_SHEET Then GoTo NextSheet

        srcLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        srcLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        If srcLastRow < 1 Or srcLastCol < 1 Then GoTo NextSheet

        Dim startRow As Long
        If isFirst Then
            startRow = 1
            isFirst = False
        Else
            startRow = 2
        End If

        For i = startRow To srcLastRow
            For c = 1 To srcLastCol
                summaryWs.Cells(destRow, c).Formula = _
                    "='" & ws.Name & "'!" & ws.Cells(i, c).Address(False, False)
            Next c
            destRow = destRow + 1
        Next i

NextSheet:
    Next ws

    Application.ScreenUpdating = True
    MsgBox "已以參照公式合併完成！共 " & (destRow - 1) & " 列。", vbInformation, "完成"
End Sub
