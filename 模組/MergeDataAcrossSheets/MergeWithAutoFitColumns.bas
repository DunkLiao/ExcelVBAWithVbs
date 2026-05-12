Attribute VB_Name = "MergeWithAutoFitColumns"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithAutoFitColumns
'功能說明: 合併活頁簿內所有工作表至彙總表，並自動調整欄寬與列高
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestMergeWithAutoFitColumns()
    Call MergeWithAutoFitColumns
End Sub

' 合併所有工作表並自動調整欄寬
Sub MergeWithAutoFitColumns()
    On Error GoTo ErrorHandler

    Dim wbSrc As Workbook
    Dim wsDest As Worksheet
    Dim ws As Worksheet
    Dim destRow As Long
    Dim lastRow As Long
    Dim headerCopied As Boolean

    Set wbSrc = ActiveWorkbook
    Set wsDest = GetOrCreateSheet(wbSrc, "彙總_AutoFit")
    destRow = 1
    headerCopied = False

    Application.ScreenUpdating = False

    For Each ws In wbSrc.Worksheets
        If ws.Name = wsDest.Name Then GoTo NextSheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow < 1 Then GoTo NextSheet

        If Not headerCopied Then
            ws.Rows(1).Copy Destination:=wsDest.Rows(destRow)
            wsDest.Rows(destRow).Font.Bold = True
            destRow = destRow + 1
            headerCopied = True
            If lastRow > 1 Then
                ws.Rows("2:" & lastRow).Copy Destination:=wsDest.Rows(destRow)
                destRow = destRow + lastRow - 1
            End If
        Else
            If lastRow > 1 Then
                ws.Rows("2:" & lastRow).Copy Destination:=wsDest.Rows(destRow)
                destRow = destRow + lastRow - 1
            End If
        End If

NextSheet:
    Next ws

    ' 自動調整欄寬與列高
    wsDest.Columns.AutoFit
    wsDest.Rows.AutoFit

    Application.ScreenUpdating = True
    MsgBox "合併完成，欄寬與列高已自動調整！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
