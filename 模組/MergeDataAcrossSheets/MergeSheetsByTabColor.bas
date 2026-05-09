Attribute VB_Name = "MergeSheetsByTabColor"
Option Explicit

' ============================================================
' 範例：合併與目前使用中工作表相同標籤顏色的工作表
' 功能：偵測目前工作表的標籤顏色，合併相同顏色的工作表
' 使用方式：先將要合併的工作表設定相同標籤顏色，再執行
' ============================================================
Sub MergeSheetsByTabColor()
    Dim wsActive    As Worksheet
    Dim ws          As Worksheet
    Dim wsSummary   As Worksheet
    Dim lngColor    As Long
    Dim lngDestRow  As Long
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim lngCopyStart As Long
    Dim blnFirst    As Boolean

    On Error GoTo ErrHandler

    Set wsActive = ActiveSheet
    lngColor = wsActive.Tab.Color

    If wsActive.Tab.ColorIndex = xlColorIndexNone Then
        MsgBox "目前工作表沒有設定標籤顏色，請先設定後再執行。", vbExclamation
        Exit Sub
    End If

    Set wsSummary = GetOrCreateSheet_TabColor(ThisWorkbook, "標籤顏色合併")
    lngDestRow = 1
    blnFirst = True
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSummary.Name Then
            If ws.Tab.ColorIndex <> xlColorIndexNone Then
                If ws.Tab.Color = lngColor Then
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
            End If
        End If
    Next ws

    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If blnFirst Then
        MsgBox "找不到相同標籤顏色的工作表。", vbExclamation
    Else
        MsgBox "標籤顏色合併完成！資料已寫入工作表：標籤顏色合併", vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet_TabColor(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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
    Set GetOrCreateSheet_TabColor = ws
End Function
