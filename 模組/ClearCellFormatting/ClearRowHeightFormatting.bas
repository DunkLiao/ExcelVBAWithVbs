Attribute VB_Name = "ClearRowHeightFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearRowHeightFormatting
'功能說明: 清除指定範圍內所有自訂列高設定，將列高還原為自動調整狀態
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestClearRowHeightFormatting()
    Dim ws As Worksheet
    Set ws = GetOrCreateRowHeightSheet(ThisWorkbook, "清除列高格式範例")
    Call FillRowHeightSampleData(ws)
    Call ApplyCustomRowHeights(ws)
    MsgBox "已套用自訂列高，請確認後按確定以清除列高設定。", vbInformation, "提示"
    Call ClearRowHeightFormatting(ws.UsedRange)
    MsgBox "列高格式已清除並還原自動調整！", vbInformation, "完成"
End Sub

Sub ClearRowHeightFormatting(ByVal rng As Range)
    Dim row As Range
    Application.ScreenUpdating = False
    For Each row In rng.Rows
        row.UseStandardHeight = True
    Next row
    rng.Rows.AutoFit
    Application.ScreenUpdating = True
End Sub

Sub ClearAllRowHeightFormatting(ByVal ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Rows.UseStandardHeight = True
    ws.Rows.AutoFit
    Application.ScreenUpdating = True
    MsgBox "工作表『" & ws.Name & "』的所有列高已還原自動調整。", _
           vbInformation, "完成"
End Sub

Private Sub ApplyCustomRowHeights(ByVal ws As Worksheet)
    ws.Rows(2).RowHeight = 40
    ws.Rows(3).RowHeight = 60
    ws.Rows(4).RowHeight = 25
    ws.Rows(5).RowHeight = 80
End Sub

Private Sub FillRowHeightSampleData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("項目", "說明")
    ws.Range("A2:B2").Value = Array("項目一", "這是第一列，列高設為40")
    ws.Range("A3:B3").Value = Array("項目二", "這是第二列，列高設為60")
    ws.Range("A4:B4").Value = Array("項目三", "這是第三列，列高設為25")
    ws.Range("A5:B5").Value = Array("項目四", "這是第四列，列高設為80")
    ws.Columns("A:B").AutoFit
End Sub

Private Function GetOrCreateRowHeightSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateRowHeightSheet = ws
End Function
