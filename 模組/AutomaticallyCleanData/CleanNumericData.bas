Attribute VB_Name = "CleanNumericData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanNumericData
'功能說明: 清理數值欄位，移除貨幣符號（NT$、$）、千分位逗號與多餘空白，轉換為數字格式
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

Sub TestCleanNumericData()
    Dim ws As Worksheet
    Set ws = GetOrCreateNumericSheet(ThisWorkbook, "數值清理範例")
    Call FillDirtyNumericData(ws)
    Call CleanNumericColumns(ws, 2, 4)
    MsgBox "數值欄位清理完成！", vbInformation, "完成"
End Sub

' 清理指定欄範圍的數值格式
Sub CleanNumericColumns(ByVal ws As Worksheet, ByVal startCol As Long, ByVal endCol As Long)
    Dim lastRow As Long
    Dim r       As Long
    Dim c       As Long
    Dim strVal  As String

    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For c = startCol To endCol
        For r = 2 To lastRow
            strVal = Trim(CStr(ws.Cells(r, c).Value))
            strVal = Replace(strVal, "NT$", "")
            strVal = Replace(strVal, "$", "")
            strVal = Replace(strVal, ",", "")
            strVal = Trim(strVal)
            If IsNumeric(strVal) And strVal <> "" Then
                ws.Cells(r, c).Value = CDbl(strVal)
                ws.Cells(r, c).NumberFormat = "#,##0.00"
            End If
        Next r
    Next c

    Application.ScreenUpdating = True
End Sub

' 填入含污染數值的測試資料
Private Sub FillDirtyNumericData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("品名", "售價", "成本", "毛利")
    ws.Range("A2:D2").Value = Array("產品A", "NT$1,200", "$800", "  400  ")
    ws.Range("A3:D3").Value = Array("產品B", "2,500.50", "NT$1,800", "$700.50")
    ws.Range("A4:D4").Value = Array("產品C", " $350 ", "200", "150")
    ws.Columns("A:D").AutoFit
End Sub

Private Function GetOrCreateNumericSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateNumericSheet = ws
End Function