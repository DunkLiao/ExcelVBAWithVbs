Attribute VB_Name = "MergeFilteredRows"
Option Explicit

' ============================================================
' 範例：合併所有工作表中，指定欄位值大於門檻值的列
' 功能：讓使用者輸入欄號與數值門檻，只複製符合條件的列
' ============================================================
Sub MergeFilteredRows()
    Dim ws          As Worksheet
    Dim wsSummary   As Worksheet
    Dim lngDestRow  As Long
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim lngRow      As Long
    Dim lngFilterCol As Long
    Dim dblThreshold As Double
    Dim blnFirst    As Boolean
    Dim strInput    As String
    Dim lngCount    As Long

    On Error GoTo ErrHandler

    strInput = InputBox("請輸入篩選欄號（例如：3 代表第 C 欄）：", "條件合併", "3")
    If strInput = "" Then Exit Sub
    lngFilterCol = CLng(strInput)

    strInput = InputBox("請輸入數值門檻（大於此值的列才合併）：", "條件合併", "0")
    If strInput = "" Then Exit Sub
    dblThreshold = CDbl(strInput)

    Set wsSummary = GetOrCreateSheet_Filter(ThisWorkbook, "條件篩選合併")
    lngDestRow = 1
    blnFirst = True
    lngCount = 0
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSummary.Name Then
            lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lngLastRow >= 2 And lngLastCol >= lngFilterCol Then
                If blnFirst Then
                    ws.Range(ws.Cells(1, 1), ws.Cells(1, lngLastCol)).Copy _
                        Destination:=wsSummary.Cells(1, 1)
                    lngDestRow = 2
                    blnFirst = False
                End If
                For lngRow = 2 To lngLastRow
                    If IsNumeric(ws.Cells(lngRow, lngFilterCol).Value) Then
                        If CDbl(ws.Cells(lngRow, lngFilterCol).Value) > dblThreshold Then
                            ws.Range(ws.Cells(lngRow, 1), ws.Cells(lngRow, lngLastCol)).Copy _
                                Destination:=wsSummary.Cells(lngDestRow, 1)
                            lngDestRow = lngDestRow + 1
                            lngCount = lngCount + 1
                        End If
                    End If
                Next lngRow
            End If
        End If
    Next ws

    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "找不到符合條件的資料。", vbExclamation
    Else
        MsgBox "條件篩選合併完成！共合併 " & lngCount & " 筆資料。", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_Filter(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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
    Set GetOrCreateSheet_Filter = ws
End Function
