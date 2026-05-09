Attribute VB_Name = "MergeDeduplicatedRows"
Option Explicit

' ============================================================
' 範例：合併所有工作表後，依第一欄去除完全重複的列
' 功能：使用 Collection 追蹤已出現的第一欄值，略過重複列
' ============================================================
Sub MergeDeduplicatedRows()
    Dim ws          As Worksheet
    Dim wsSummary   As Worksheet
    Dim lngDestRow  As Long
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim lngRow      As Long
    Dim blnFirst    As Boolean
    Dim colKeys     As Collection
    Dim strKey      As String
    Dim blnDup      As Boolean

    On Error GoTo ErrHandler

    Set colKeys = New Collection
    Set wsSummary = GetOrCreateSheet_Dedup(ThisWorkbook, "去重複合併")
    lngDestRow = 1
    blnFirst = True
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSummary.Name Then
            lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lngLastRow >= 1 And lngLastCol >= 1 Then
                If blnFirst Then
                    ' 複製標題列（不去重）
                    ws.Range(ws.Cells(1, 1), ws.Cells(1, lngLastCol)).Copy _
                        Destination:=wsSummary.Cells(1, 1)
                    lngDestRow = 2
                    blnFirst = False
                End If
                For lngRow = 2 To lngLastRow
                    strKey = CStr(ws.Cells(lngRow, 1).Value)
                    blnDup = False
                    On Error Resume Next
                    colKeys.Add strKey, strKey
                    If Err.Number <> 0 Then blnDup = True
                    On Error GoTo ErrHandler
                    If Not blnDup Then
                        ws.Range(ws.Cells(lngRow, 1), ws.Cells(lngRow, lngLastCol)).Copy _
                            Destination:=wsSummary.Cells(lngDestRow, 1)
                        lngDestRow = lngDestRow + 1
                    End If
                Next lngRow
            End If
        End If
    Next ws

    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "所有工作表均無資料可合併。", vbExclamation
    Else
        MsgBox "去除重複合併完成！共寫入 " & (lngDestRow - 2) & " 筆不重複資料。", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_Dedup(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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
    Set GetOrCreateSheet_Dedup = ws
End Function
