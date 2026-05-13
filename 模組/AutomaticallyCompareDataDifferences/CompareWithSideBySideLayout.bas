Attribute VB_Name = "CompareWithSideBySideLayout"
Option Explicit
'*************************************************************************************
'模組名稱: CompareWithSideBySideLayout
'功能說明: 將兩張工作表的資料並排顯示於同一工作表，並標示差異儲存格
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub CompareSheetsSideBySide()
    On Error GoTo ErrHandler
    Dim wb        As Workbook
    Dim ws1       As Worksheet
    Dim ws2       As Worksheet
    Dim wsOut     As Worksheet
    Dim wsTemp    As Worksheet
    Dim r         As Long
    Dim c         As Long
    Dim lastRow   As Long
    Dim lastCol   As Long
    Dim diffCount As Long
    Dim idx       As Long
    Dim offset    As Long
    Dim v1        As String
    Dim v2        As String

    Set wb = ThisWorkbook
    idx = 0
    For Each wsTemp In wb.Worksheets
        If wsTemp.Visible = xlSheetVisible Then
            idx = idx + 1
            If idx = 1 Then Set ws1 = wsTemp
            If idx = 2 Then Set ws2 = wsTemp
        End If
        If idx = 2 Then Exit For
    Next wsTemp

    If ws1 Is Nothing Or ws2 Is Nothing Then
        MsgBox "活頁簿中可見工作表不足兩張，無法進行比對。", vbExclamation, "提示"
        Exit Sub
    End If
    Set wsOut = GetOrCreateSheetSBS(wb, "並排比對結果")

    lastRow = Application.WorksheetFunction.Max( _
        ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row, _
        ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row)
    lastCol = Application.WorksheetFunction.Max( _
        ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column, _
        ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column)

    wsOut.Range("A1").Value = "來源工作表：" & ws1.Name
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Color = RGB(0, 70, 127)
    offset = lastCol + 2
    wsOut.Cells(1, offset).Value = "比對工作表：" & ws2.Name
    wsOut.Cells(1, offset).Font.Bold = True
    wsOut.Cells(1, offset).Font.Color = RGB(127, 0, 0)
    diffCount = 0
    For r = 1 To lastRow
        For c = 1 To lastCol
            v1 = CStr(ws1.Cells(r, c).Value)
            v2 = CStr(ws2.Cells(r, c).Value)
            wsOut.Cells(r + 1, c).Value = v1
            wsOut.Cells(r + 1, c + offset - 1).Value = v2
            If v1 <> v2 Then
                wsOut.Cells(r + 1, c).Interior.Color = RGB(255, 199, 206)
                wsOut.Cells(r + 1, c + offset - 1).Interior.Color = RGB(255, 199, 206)
                diffCount = diffCount + 1
            End If
        Next c
    Next r
    wsOut.Columns.AutoFit
    wsOut.Activate
    MsgBox "並排比對完成！共發現 " & diffCount & " 處差異（紅色標示）。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetSBS(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetSBS = ws
End Function

