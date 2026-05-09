Attribute VB_Name = "MergeBySheetNameList"
Option Explicit

' ============================================================
' 範例：依使用者在儲存格範圍中填入的工作表名稱清單來合併
' 功能：讀取指定欄位的工作表名稱清單，依序合併對應工作表
' 使用方式：在任一工作表的 A 欄列出要合併的工作表名稱後執行
' ============================================================
Sub MergeBySheetNameList()
    Dim wsConfig    As Worksheet
    Dim ws          As Worksheet
    Dim wsSummary   As Worksheet
    Dim lngDestRow  As Long
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim lngCopyStart As Long
    Dim lngListRow  As Long
    Dim lngListEnd  As Long
    Dim blnFirst    As Boolean
    Dim strSheetName As String
    Dim strConfigSheet As String
    Dim intCount    As Integer

    On Error GoTo ErrHandler

    strConfigSheet = InputBox("請輸入包含工作表名稱清單的工作表名稱：", "名稱清單合併", ActiveSheet.Name)
    If strConfigSheet = "" Then Exit Sub

    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets(strConfigSheet)
    On Error GoTo ErrHandler
    If wsConfig Is Nothing Then
        MsgBox "找不到工作表：" & strConfigSheet, vbExclamation
        Exit Sub
    End If

    lngListEnd = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    If lngListEnd < 1 Then
        MsgBox "A 欄未填入任何工作表名稱。", vbExclamation
        Exit Sub
    End If

    Set wsSummary = GetOrCreateSheet_List(ThisWorkbook, "名稱清單合併")
    lngDestRow = 1
    blnFirst = True
    intCount = 0
    Application.ScreenUpdating = False

    For lngListRow = 1 To lngListEnd
        strSheetName = Trim(CStr(wsConfig.Cells(lngListRow, 1).Value))
        If strSheetName <> "" And strSheetName <> wsSummary.Name Then
            Set ws = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(strSheetName)
            On Error GoTo ErrHandler
            If Not ws Is Nothing Then
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
                        intCount = intCount + 1
                    End If
                End If
            End If
        End If
    Next lngListRow

    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "名稱清單中的工作表均無資料可合併。", vbExclamation
    Else
        MsgBox "名稱清單合併完成！共合併 " & intCount & " 個工作表。", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_List(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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
    Set GetOrCreateSheet_List = ws
End Function
