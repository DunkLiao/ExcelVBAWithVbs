Attribute VB_Name = "MergeSortedResult"
Option Explicit

' ============================================================
' 範例：合併所有工作表後，依使用者指定的欄位遞增排序
' 功能：完成合併後，對結果工作表執行 Sort 操作
' ============================================================
Sub MergeSortedResult()
    Dim ws          As Worksheet
    Dim wsSummary   As Worksheet
    Dim lngDestRow  As Long
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim lngSortCol  As Long
    Dim blnFirst    As Boolean
    Dim strInput    As String
    Dim rngData     As Range

    On Error GoTo ErrHandler

    strInput = InputBox("合併後依哪一欄排序？請輸入欄號（例如：1 代表第 A 欄）：", "排序合併", "1")
    If strInput = "" Then Exit Sub
    lngSortCol = CLng(strInput)
    If lngSortCol < 1 Then
        MsgBox "欄號無效，請輸入大於 0 的數值。", vbExclamation
        Exit Sub
    End If

    Set wsSummary = GetOrCreateSheet_Sort(ThisWorkbook, "排序後合併")
    lngDestRow = 1
    blnFirst = True
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSummary.Name Then
            lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lngLastRow >= 1 And lngLastCol >= 1 Then
                If blnFirst Then
                    ws.Range(ws.Cells(1, 1), ws.Cells(lngLastRow, lngLastCol)).Copy _
                        Destination:=wsSummary.Cells(lngDestRow, 1)
                    lngDestRow = lngDestRow + lngLastRow
                    blnFirst = False
                Else
                    If lngLastRow >= 2 Then
                        ws.Range(ws.Cells(2, 1), ws.Cells(lngLastRow, lngLastCol)).Copy _
                            Destination:=wsSummary.Cells(lngDestRow, 1)
                        lngDestRow = lngDestRow + lngLastRow - 1
                    End If
                End If
            End If
        End If
    Next ws

    ' 對合併結果排序（保留標題列）
    If lngDestRow > 2 Then
        lngLastCol = wsSummary.Cells(1, wsSummary.Columns.Count).End(xlToLeft).Column
        Set rngData = wsSummary.Range(wsSummary.Cells(1, 1), wsSummary.Cells(lngDestRow - 1, lngLastCol))
        rngData.Sort Key1:=wsSummary.Cells(2, lngSortCol), _
                      Order1:=xlAscending, _
                      Header:=xlYes
    End If

    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "所有工作表均無資料可合併。", vbExclamation
    Else
        MsgBox "合併並排序完成！結果已依第 " & lngSortCol & " 欄遞增排序。", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_Sort(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    Set GetOrCreateSheet_Sort = ws
End Function
