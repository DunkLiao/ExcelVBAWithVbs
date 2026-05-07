Attribute VB_Name = "CleanDuplicateAndBlank"
Option Explicit
'*************************************************************************************
'模組名稱: CleanDuplicateAndBlank
'功能說明: 自動清理資料：刪除重複列、移除空白列及修整多餘空格
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口（建立含髒資料的範例後自動清理）
Sub TestCleanData()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "待清理資料")
    Call FillDirtyData(ws)
    Call CleanAllData(ws)
End Sub

' 執行完整資料清理（整合所有清理步驟）
Sub CleanAllData(ByVal ws As Worksheet)
    Application.ScreenUpdating = False

    Call TrimAllStringCells(ws)
    Call RemoveBlankRows(ws)
    Call RemoveDuplicateRows(ws)

    Application.ScreenUpdating = True

    ws.Activate
    MsgBox "資料清理完成！已執行：修整空格 -> 刪除空白列 -> 刪除重複列", _
           vbInformation, "完成"
End Sub

' 修整工作表內所有字串儲存格的多餘空格
Private Sub TrimAllStringCells(ByVal ws As Worksheet)
    Dim cell As Range
    Dim usedRng As Range

    Set usedRng = ws.UsedRange
    For Each cell In usedRng
        If cell.Value <> "" And VarType(cell.Value) = vbString Then
            cell.Value = Trim(cell.Value)
        End If
    Next cell
End Sub

' 刪除工作表內的整列空白列
Private Sub RemoveBlankRows(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 從最後一列往上刪，避免刪除後列索引偏移
    For r = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

' 刪除工作表內的重複列（保留第一筆）
Private Sub RemoveDuplicateRows(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim colArray() As Integer
    Dim i As Integer

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then Exit Sub

    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ReDim colArray(1 To lastCol)
    For i = 1 To lastCol
        colArray(i) = i
    Next i

    dataRange.RemoveDuplicates Columns:=colArray, Header:=xlYes
End Sub

' 填入含有髒資料的範例（含空白列、重複列、多餘空格）
Private Sub FillDirtyData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "電話"
    ws.Range("A2").Value = "  張三  "
    ws.Range("B2").Value = "研發部"
    ws.Range("C2").Value = "0912-345678"
    ws.Range("A3").Value = ""
    ws.Range("B3").Value = ""
    ws.Range("C3").Value = ""
    ws.Range("A4").Value = "李四"
    ws.Range("B4").Value = "  業務部  "
    ws.Range("C4").Value = "0923-456789"
    ws.Range("A5").Value = "張三"
    ws.Range("B5").Value = "研發部"
    ws.Range("C5").Value = "0912-345678"
    ws.Range("A6").Value = "王五"
    ws.Range("B6").Value = "行政部"
    ws.Range("C6").Value = "0934-567890"
    ws.Range("A7").Value = ""
    ws.Range("B7").Value = ""
    ws.Range("C7").Value = ""
    ws.Range("A8").Value = "李四"
    ws.Range("B8").Value = "業務部"
    ws.Range("C8").Value = "0923-456789"
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表，並清除內容
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
