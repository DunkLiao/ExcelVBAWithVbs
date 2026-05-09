Attribute VB_Name = "FillMissingValues"
Option Explicit
'*************************************************************************************
'模組名稱: FillMissingValues
'功能說明: 向下填補空白儲存格（Fill Down），適用於合併儲存格來源或報表遺漏值填補
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

Sub TestFillMissingValues()
    Dim ws As Worksheet
    Set ws = GetOrCreateFillSheet(ThisWorkbook, "遺漏值填補範例")
    Call FillSparseData(ws)
    Call FillDownColumn(ws, 1)
    Call FillDownColumn(ws, 2)
    ws.Columns("A:D").AutoFit
    MsgBox "遺漏值向下填補完成！", vbInformation, "完成"
End Sub

' 對指定欄進行向下填補（Fill Down）
Sub FillDownColumn(ByVal ws As Worksheet, ByVal colIndex As Long)
    Dim lastRow  As Long
    Dim r        As Long
    Dim lastVal  As Variant

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    lastVal = ""

    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, colIndex).Value)) = "" Then
            ' 空白儲存格：填入上一個非空值
            If lastVal <> "" Then
                ws.Cells(r, colIndex).Value = lastVal
            End If
        Else
            ' 非空儲存格：更新記錄值
            lastVal = ws.Cells(r, colIndex).Value
        End If
    Next r
End Sub

' 對工作表中多個欄位批次執行向下填補
Sub FillDownMultipleColumns(ByVal ws As Worksheet, ParamArray colIndexes() As Variant)
    Dim idx As Variant
    For Each idx In colIndexes
        Call FillDownColumn(ws, CLng(idx))
    Next idx
End Sub

Private Sub FillSparseData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("部門", "組別", "員工姓名", "職稱")
    ws.Range("A2:D2").Value = Array("業務部", "北區業務組", "陳大文", "業務員")
    ws.Range("A3:D3").Value = Array("", "", "林小美", "資深業務")
    ws.Range("A4:D4").Value = Array("", "南區業務組", "王志強", "業務員")
    ws.Range("A5:D5").Value = Array("", "", "張明宏", "業務員")
    ws.Range("A6:D6").Value = Array("技術部", "研發組", "李俊賢", "工程師")
    ws.Range("A7:D7").Value = Array("", "", "劉雅婷", "資深工程師")
    ws.Range("A8:D8").Value = Array("", "測試組", "吳建志", "測試工程師")
    ws.Columns("A:D").AutoFit
End Sub

Private Function GetOrCreateFillSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateFillSheet = ws
End Function