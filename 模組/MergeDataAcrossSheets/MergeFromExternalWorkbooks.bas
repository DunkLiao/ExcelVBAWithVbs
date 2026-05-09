Attribute VB_Name = "MergeFromExternalWorkbooks"
Option Explicit

' ============================================================
' 範例：從指定資料夾中所有 .xlsx/.xls 活頁簿合併第一張工作表
' 功能：開啟每個活頁簿，複製第一張工作表資料，合併後關閉
' ============================================================
Sub MergeFromExternalWorkbooks()
    Dim strFolder   As String
    Dim strFile     As String
    Dim wbSrc       As Workbook
    Dim wsSrc       As Worksheet
    Dim wsSummary   As Worksheet
    Dim lngDestRow  As Long
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim lngCopyStart As Long
    Dim blnFirst    As Boolean
    Dim intCount    As Integer

    On Error GoTo ErrHandler

    ' 讓使用者選取資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選取包含活頁簿的資料夾"
        If .Show <> -1 Then Exit Sub
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "" Then strFolder = strFolder & ""

    Set wsSummary = GetOrCreateSheet_Ext(ThisWorkbook, "外部活頁簿合併")
    lngDestRow = 1
    blnFirst = True
    intCount = 0
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    strFile = Dir(strFolder & "*.xls*")
    Do While strFile <> ""
        If strFolder & strFile <> ThisWorkbook.FullName Then
            Set wbSrc = Workbooks.Open(strFolder & strFile, ReadOnly:=True)
            Set wsSrc = wbSrc.Worksheets(1)
            lngLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            lngLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
            If lngLastRow >= 1 And lngLastCol >= 1 Then
                If blnFirst Then
                    lngCopyStart = 1
                    blnFirst = False
                Else
                    lngCopyStart = 2  ' 略過標題列
                End If
                If lngLastRow >= lngCopyStart Then
                    wsSrc.Range(wsSrc.Cells(lngCopyStart, 1), wsSrc.Cells(lngLastRow, lngLastCol)).Copy _
                        Destination:=wsSummary.Cells(lngDestRow, 1)
                    lngDestRow = lngDestRow + (lngLastRow - lngCopyStart + 1)
                    intCount = intCount + 1
                End If
            End If
            wbSrc.Close SaveChanges:=False
        End If
        strFile = Dir
    Loop

    wsSummary.Columns.AutoFit
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If intCount = 0 Then
        MsgBox "資料夾內找不到可合併的活頁簿。", vbExclamation
    Else
        MsgBox "已合併 " & intCount & " 個活頁簿的資料！", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_Ext(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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
    Set GetOrCreateSheet_Ext = ws
End Function
