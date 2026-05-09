Attribute VB_Name = "MergeSheetsHorizontally"
Option Explicit

' ============================================================
' 範例：橫向合併工作表（每張工作表的資料並排於同一列）
' 功能：將各工作表資料由左至右依序貼至同一工作表，以欄分隔
' 注意：各工作表列數可不同，以最多列者為準
' ============================================================
Sub MergeSheetsHorizontally()
    Dim ws          As Worksheet
    Dim wsSummary   As Worksheet
    Dim lngDestCol  As Long
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim blnFirst    As Boolean

    On Error GoTo ErrHandler

    Set wsSummary = GetOrCreateSheet_Horiz(ThisWorkbook, "橫向並排合併")
    lngDestCol = 1
    blnFirst = True
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSummary.Name Then
            lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lngLastRow >= 1 And lngLastCol >= 1 Then
                ws.Range(ws.Cells(1, 1), ws.Cells(lngLastRow, lngLastCol)).Copy _
                    Destination:=wsSummary.Cells(1, lngDestCol)
                If blnFirst Then
                    blnFirst = False
                Else
                    ' 插入空白分隔欄
                    lngDestCol = lngDestCol
                End If
                lngDestCol = lngDestCol + lngLastCol + 1
            End If
        End If
    Next ws

    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "所有工作表均無資料可合併。", vbExclamation
    Else
        MsgBox "橫向並排合併完成！各工作表資料已由左至右排列。", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_Horiz(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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
    Set GetOrCreateSheet_Horiz = ws
End Function
