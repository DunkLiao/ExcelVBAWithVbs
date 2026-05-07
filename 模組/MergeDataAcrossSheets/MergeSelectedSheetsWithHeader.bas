Attribute VB_Name = "MergeSelectedSheetsWithHeader"
Option Explicit

' ============================================================
' 範例：合併使用者選定的工作表資料至彙總表（自動處理標題列）
' 功能：僅合併名稱符合前置字串的工作表，第一個工作表保留標題
' ============================================================
Sub MergeSelectedSheetsWithHeader()
    Dim ws          As Worksheet
    Dim wsSummary   As Worksheet
    Dim strPrefix   As String
    Dim lngDestRow  As Long
    Dim lngLastRow  As Long
    Dim lngStart    As Long
    Dim blnFirst    As Boolean

    ' --- 詢問工作表名稱前置字串 ---
    strPrefix = InputBox( _
        "請輸入要合併的工作表名稱前置字串" & Chr(10) & _
        "（例如：輸入 ""月報"" 將合併所有以月報開頭的工作表）", _
        "合併工作表", "Sheet")

    If strPrefix = "" Then
        MsgBox "未輸入前置字串，操作取消。", vbInformation
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ' --- 建立或清除彙總工作表 ---
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Worksheets("合併彙總")
    On Error GoTo ErrHandler

    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSummary.Name = "合併彙總"
    Else
        wsSummary.Cells.Clear
    End If

    lngDestRow = 1
    blnFirst = True
    Application.ScreenUpdating = False

    ' --- 依前置字串篩選並合併 ---
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSummary.Name Then
            If Left(ws.Name, Len(strPrefix)) = strPrefix Then
                lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                If lngLastRow >= 1 Then
                    If blnFirst Then
                        lngStart = 1
                        blnFirst = False
                    Else
                        lngStart = 2  ' 跳過標題列
                    End If
                    ws.Rows(lngStart & ":" & lngLastRow).Copy _
                        Destination:=wsSummary.Cells(lngDestRow, 1)
                    lngDestRow = lngDestRow + (lngLastRow - lngStart + 1)
                End If
            End If
        End If
    Next ws

    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "找不到符合前置字串「" & strPrefix & "」的工作表。", vbExclamation
    Else
        MsgBox "合併完成！資料已寫入工作表：合併彙總", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
