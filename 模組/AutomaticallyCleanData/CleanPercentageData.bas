Attribute VB_Name = "CleanPercentageData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanPercentageData
'功能說明: 清理百分比欄位，移除 % 符號並轉換為小數（如 85% -> 0.85），統一儲存格格式
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

Sub TestCleanPercentageData()
    Dim ws As Worksheet
    Set ws = GetOrCreatePctSheet(ThisWorkbook, "百分比清理範例")
    Call FillDirtyPercentageData(ws)
    Call CleanPercentageColumns(ws, 2, 4)
    ws.Columns("A:D").AutoFit
    MsgBox "百分比欄位清理完成！", vbInformation, "完成"
End Sub

' 清理指定欄範圍的百分比資料，轉換為 Excel 百分比格式
Sub CleanPercentageColumns(ByVal ws As Worksheet, ByVal startCol As Long, ByVal endCol As Long)
    Dim lastRow  As Long
    Dim r        As Long
    Dim c        As Long
    Dim strVal   As String
    Dim dblVal   As Double

    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For c = startCol To endCol
        For r = 2 To lastRow
            strVal = Trim(CStr(ws.Cells(r, c).Value))
            If strVal <> "" Then
                ' 移除 % 符號、空白及千分位逗號
                strVal = Replace(strVal, "%", "")
                strVal = Replace(strVal, ",", "")
                strVal = Trim(strVal)
                If IsNumeric(strVal) Then
                    dblVal = CDbl(strVal)
                    ' 若數值 > 1，視為百分比整數（如 85.5 代表 85.5%），除以 100
                    If dblVal > 1 Then dblVal = dblVal / 100
                    ws.Cells(r, c).Value = dblVal
                    ws.Cells(r, c).NumberFormat = "0.00%"
                End If
            End If
        Next r
    Next c

    Application.ScreenUpdating = True
End Sub

Private Sub FillDirtyPercentageData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("姓名", "業績達成率", "客戶滿意度", "出席率")
    ws.Range("A2:D2").Value = Array("陳大文", "85%", "92.5%", "100%")
    ws.Range("A3:D3").Value = Array("林小美", " 78.3% ", "88%", " 95%")
    ws.Range("A4:D4").Value = Array("王志強", "105.2%", "76%", "90%")
    ws.Range("A5:D5").Value = Array("張明宏", "0.92", "0.85", "1.00")
    ws.Columns("A:D").AutoFit
End Sub

Private Function GetOrCreatePctSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreatePctSheet = ws
End Function