Attribute VB_Name = "CleanDateFormat"
Option Explicit
'*************************************************************************************
'模組名稱: CleanDateFormat
'功能說明: 統一日期格式，將民國年、各種分隔符的日期轉換為 yyyy/mm/dd 標準格式
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

Sub TestCleanDateFormat()
    Dim ws As Worksheet
    Set ws = GetOrCreateDateSheet(ThisWorkbook, "日期清理範例")
    Call FillDirtyDateData(ws)
    Call StandardizeDateColumn(ws, 2)
    MsgBox "日期欄位格式統一完成！", vbInformation, "完成"
End Sub

' 將指定欄位的日期統一轉換為 yyyy/mm/dd 格式
Sub StandardizeDateColumn(ByVal ws As Worksheet, ByVal colIndex As Long)
    Dim lastRow As Long
    Dim r       As Long
    Dim strVal  As String
    Dim dtVal   As Date
    Dim parts() As String
    Dim yearNum As Long

    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        strVal = Trim(CStr(ws.Cells(r, colIndex).Value))
        If strVal <> "" Then
            ' 統一分隔符為 /
            strVal = Replace(strVal, "-", "/")
            strVal = Replace(strVal, ".", "/")
            ' 判斷是否為民國年（3位數且 < 200）
            parts = Split(strVal, "/")
            If UBound(parts) = 2 Then
                yearNum = 0
                If IsNumeric(parts(0)) Then yearNum = CLng(parts(0))
                If yearNum > 0 And yearNum < 200 Then
                    ' 民國年加 1911 轉為西元年
                    strVal = CStr(yearNum + 1911) & "/" & parts(1) & "/" & parts(2)
                End If
            End If
            If IsDate(strVal) Then
                dtVal = CDate(strVal)
                ws.Cells(r, colIndex).Value = dtVal
                ws.Cells(r, colIndex).NumberFormat = "yyyy/mm/dd"
            End If
        End If
    Next r

    Application.ScreenUpdating = True
End Sub

Private Sub FillDirtyDateData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("編號", "日期原始值", "備註")
    ws.Range("A2:C2").Value = Array(1, "2025-01-15", "西元年橫線")
    ws.Range("A3:C3").Value = Array(2, "114/05/09", "民國年斜線")
    ws.Range("A4:C4").Value = Array(3, "2024.12.31", "西元年點號")
    ws.Range("A5:C5").Value = Array(4, "113/07/20", "民國年")
    ws.Range("A6:C6").Value = Array(5, "2026/03/01", "標準格式")
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateDateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateDateSheet = ws
End Function