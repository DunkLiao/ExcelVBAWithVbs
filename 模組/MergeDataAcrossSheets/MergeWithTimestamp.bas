Attribute VB_Name = "MergeWithTimestamp"
Option Explicit

' ============================================================
' 範例：合併所有工作表並在每列新增來源工作表名稱與合併時間
' 功能：每列資料附加「來源工作表」與「合併時間」兩個欄位
' ============================================================
Sub MergeWithTimestamp()
    Dim ws          As Worksheet
    Dim wsSummary   As Worksheet
    Dim lngDestRow  As Long
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim lngRow      As Long
    Dim lngHeaderCol As Long
    Dim blnFirst    As Boolean
    Dim dtMerge     As Date

    On Error GoTo ErrHandler

    dtMerge = Now
    Set wsSummary = GetOrCreateSheet_TS(ThisWorkbook, "時間戳記合併")
    lngDestRow = 1
    blnFirst = True
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSummary.Name Then
            lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lngLastRow >= 1 And lngLastCol >= 1 Then
                If blnFirst Then
                    ' 複製標題列並加入兩個附加欄標題
                    ws.Range(ws.Cells(1, 1), ws.Cells(1, lngLastCol)).Copy _
                        Destination:=wsSummary.Cells(1, 1)
                    lngHeaderCol = lngLastCol + 1
                    wsSummary.Cells(1, lngHeaderCol).Value = "來源工作表"
                    wsSummary.Cells(1, lngHeaderCol + 1).Value = "合併時間"
                    lngDestRow = 2
                    blnFirst = False
                End If
                If lngLastRow >= 2 Then
                    ws.Range(ws.Cells(2, 1), ws.Cells(lngLastRow, lngLastCol)).Copy _
                        Destination:=wsSummary.Cells(lngDestRow, 1)
                    For lngRow = lngDestRow To lngDestRow + lngLastRow - 2
                        wsSummary.Cells(lngRow, lngLastCol + 1).Value = ws.Name
                        wsSummary.Cells(lngRow, lngLastCol + 2).Value = dtMerge
                    Next lngRow
                    lngDestRow = lngDestRow + lngLastRow - 1
                End If
            End If
        End If
    Next ws

    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "所有工作表均無資料可合併。", vbExclamation
    Else
        MsgBox "時間戳記合併完成！合併時間：" & Format(dtMerge, "yyyy/mm/dd hh:mm:ss"), vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_TS(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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
    Set GetOrCreateSheet_TS = ws
End Function
