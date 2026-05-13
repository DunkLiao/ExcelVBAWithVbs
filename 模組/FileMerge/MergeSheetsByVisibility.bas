Attribute VB_Name = "MergeSheetsByVisibility"
Option Explicit
'*************************************************************************************
'模組名稱: MergeSheetsByVisibility
'功能說明: 將活頁簿中所有可見工作表的資料合併至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub MergeAllVisibleSheetsData()
    On Error GoTo ErrHandler
    Dim wbSrc   As Workbook
    Dim wsDest  As Worksheet
    Dim ws2     As Worksheet
    Dim destRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim isFirst As Boolean

    Set wbSrc = ThisWorkbook
    Set wsDest = GetOrCreateSheetVis(wbSrc, "可見工作表合併結果")
    destRow = 1
    isFirst = True

    For Each ws2 In wbSrc.Worksheets
        If ws2.Visible = xlSheetVisible And ws2.Name <> wsDest.Name Then
            lastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
            lastCol = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column
            If lastRow >= 1 And lastCol >= 1 Then
                If isFirst Then
                    ws2.Range(ws2.Cells(1, 1), ws2.Cells(1, lastCol)).Copy _
                        Destination:=wsDest.Cells(destRow, 1)
                    destRow = destRow + 1
                    isFirst = False
                End If
                If lastRow >= 2 Then
                    ws2.Range(ws2.Cells(2, 1), ws2.Cells(lastRow, lastCol)).Copy _
                        Destination:=wsDest.Cells(destRow, 1)
                    destRow = destRow + lastRow - 1
                End If
            End If
        End If
    Next ws2

    wsDest.Columns.AutoFit
    wsDest.Activate
    MsgBox "所有可見工作表已合併完成，共 " & destRow - 2 & " 列資料。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetVis(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetVis = ws
End Function

