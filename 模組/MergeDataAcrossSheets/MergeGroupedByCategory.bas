Attribute VB_Name = "MergeGroupedByCategory"
Option Explicit

' ============================================================
' 範例：依指定欄位的分類值，將各工作表資料分組寫入不同工作表
' 功能：每個分類值對應一張目標工作表，資料分流合併
' ============================================================
Sub MergeGroupedByCategory()
    Dim ws          As Worksheet
    Dim wsTarget    As Worksheet
    Dim lngCatCol   As Long
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim lngRow      As Long
    Dim strCat      As String
    Dim strInput    As String
    Dim colTargets  As Collection
    Dim blnFirst    As Boolean
    Dim lngDestRow  As Long
    Dim strSummaryName As String
    Dim strTargetName As String

    On Error GoTo ErrHandler

    strInput = InputBox("請輸入分類欄號（例如：1 代表第 A 欄）：", "分類合併", "1")
    If strInput = "" Then Exit Sub
    lngCatCol = CLng(strInput)
    If lngCatCol < 1 Then
        MsgBox "欄號無效，請輸入大於 0 的數值。", vbExclamation
        Exit Sub
    End If

    Set colTargets = New Collection
    blnFirst = True
    strSummaryName = "分類合併_"
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, Len(strSummaryName)) <> strSummaryName Then
            lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lngLastRow >= 2 And lngLastCol >= lngCatCol Then
                For lngRow = 2 To lngLastRow
                    strCat = Trim(CStr(ws.Cells(lngRow, lngCatCol).Value))
                    If strCat <> "" Then
                        strTargetName = Left(strSummaryName & strCat, 31)
                        Set wsTarget = Nothing
                        On Error Resume Next
                        Set wsTarget = colTargets(strTargetName)
                        On Error GoTo ErrHandler
                        If wsTarget Is Nothing Then
                            Set wsTarget = GetOrCreateSheet_Cat(ThisWorkbook, strTargetName)
                            ' 寫入標題列
                            ws.Range(ws.Cells(1, 1), ws.Cells(1, lngLastCol)).Copy _
                                Destination:=wsTarget.Cells(1, 1)
                            colTargets.Add wsTarget, strTargetName
                        End If
                        lngDestRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1
                        ws.Range(ws.Cells(lngRow, 1), ws.Cells(lngRow, lngLastCol)).Copy _
                            Destination:=wsTarget.Cells(lngDestRow, 1)
                        blnFirst = False
                    End If
                Next lngRow
            End If
        End If
    Next ws

    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "找不到可分類合併的資料。", vbExclamation
    Else
        MsgBox "分類合併完成！各分類資料已分別寫入對應工作表。", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_Cat(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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
    Set GetOrCreateSheet_Cat = ws
End Function
