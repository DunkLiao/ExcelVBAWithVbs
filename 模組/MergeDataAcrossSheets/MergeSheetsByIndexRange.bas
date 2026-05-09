Attribute VB_Name = "MergeSheetsByIndexRange"
Option Explicit

' ============================================================
' 範例：依工作表索引範圍合併資料至摘要工作表
' 功能：讓使用者輸入起始與結束索引，合併該範圍內的工作表
' ============================================================
Sub MergeSheetsByIndexRange()
    Dim lngStart     As Long
    Dim lngEnd       As Long
    Dim lngDestRow   As Long
    Dim lngLastRow   As Long
    Dim lngLastCol   As Long
    Dim lngIdx       As Long
    Dim lngCopyStart As Long
    Dim blnFirst     As Boolean
    Dim ws           As Worksheet
    Dim wsSummary    As Worksheet
    Dim strInput     As String

    On Error GoTo ErrHandler

    strInput = InputBox("請輸入起始工作表索引（最小值為 1）：", "合併工作表", "1")
    If strInput = "" Then Exit Sub
    lngStart = CLng(strInput)

    strInput = InputBox("請輸入結束工作表索引（最大值為 " & ThisWorkbook.Sheets.Count & "）：", "合併工作表", CStr(ThisWorkbook.Sheets.Count))
    If strInput = "" Then Exit Sub
    lngEnd = CLng(strInput)

    If lngStart < 1 Or lngEnd > ThisWorkbook.Sheets.Count Or lngStart > lngEnd Then
        MsgBox "索引範圍無效，請重新輸入。", vbExclamation
        Exit Sub
    End If

    Set wsSummary = GetOrCreateSheet_IdxRange(ThisWorkbook, "索引範圍合併")
    lngDestRow = 1
    blnFirst = True
    Application.ScreenUpdating = False

    For lngIdx = lngStart To lngEnd
        Set ws = ThisWorkbook.Sheets(lngIdx)
        If ws.Name <> wsSummary.Name Then
            lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lngLastRow >= 1 And lngLastCol >= 1 Then
                If blnFirst Then
                    lngCopyStart = 1
                    blnFirst = False
                Else
                    lngCopyStart = 2  ' 略過標題列
                End If
                If lngLastRow >= lngCopyStart Then
                    ws.Range(ws.Cells(lngCopyStart, 1), ws.Cells(lngLastRow, lngLastCol)).Copy _
                        Destination:=wsSummary.Cells(lngDestRow, 1)
                    lngDestRow = lngDestRow + (lngLastRow - lngCopyStart + 1)
                End If
            End If
        End If
    Next lngIdx

    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "指定索引範圍內無資料可合併。", vbExclamation
    Else
        MsgBox "索引範圍合併完成！資料已寫入工作表：索引範圍合併", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_IdxRange(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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
    Set GetOrCreateSheet_IdxRange = ws
End Function
