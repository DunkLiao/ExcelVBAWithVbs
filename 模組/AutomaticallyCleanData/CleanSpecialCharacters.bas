Attribute VB_Name = "CleanSpecialCharacters"
Option Explicit
'*************************************************************************************
'模組名稱: CleanSpecialCharacters
'功能說明: 移除儲存格中的特殊字元，包含 @#$%^&*!~|\ 等符號，保留中英數與空白
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

Sub TestCleanSpecialCharacters()
    Dim ws As Worksheet
    Set ws = GetOrCreateSpecialCharSheet(ThisWorkbook, "特殊字元清理範例")
    Call FillDirtySpecialCharData(ws)
    Call RemoveSpecialCharsFromRange(ws.UsedRange.Offset(1))
    ws.Columns("A:C").AutoFit
    MsgBox "特殊字元清理完成！", vbInformation, "完成"
End Sub

' 移除範圍內所有儲存格的特殊字元
Sub RemoveSpecialCharsFromRange(ByVal rng As Range)
    Dim cell       As Range
    Dim strVal     As String
    Dim cleanVal   As String
    Dim i          As Integer
    Dim charCode   As Integer
    Dim specialSet As String

    ' 定義要移除的特殊字元清單
    specialSet = "@#$%^&*!~|\/<>{}[]"
    Application.ScreenUpdating = False

    For Each cell In rng
        If Not IsEmpty(cell) And VarType(cell.Value) = vbString Then
            strVal = cell.Value
            cleanVal = ""
            For i = 1 To Len(strVal)
                charCode = Asc(Mid(strVal, i, 1))
                ' 保留中文（charCode < 0 表示雙位元組）、英數及空白
                If charCode < 0 Or (charCode >= 48 And charCode <= 57) Or _
                   (charCode >= 65 And charCode <= 90) Or _
                   (charCode >= 97 And charCode <= 122) Or charCode = 32 Then
                    cleanVal = cleanVal & Mid(strVal, i, 1)
                ElseIf InStr(specialSet, Mid(strVal, i, 1)) = 0 Then
                    cleanVal = cleanVal & Mid(strVal, i, 1)
                End If
            Next i
            If cell.Value <> cleanVal Then cell.Value = Trim(cleanVal)
        End If
    Next cell

    Application.ScreenUpdating = True
End Sub

Private Sub FillDirtySpecialCharData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("產品代碼", "描述", "備註")
    ws.Range("A2:C2").Value = Array("PRD@001", "高品質#商品", "庫存@正常!")
    ws.Range("A3:C3").Value = Array("PRD#002!", "$$特價商品$$", "注意&庫存")
    ws.Range("A4:C4").Value = Array("PRD|003", "標準^商品", "正常|供貨")
    ws.Range("A5:C5").Value = Array("PRD~004", "特殊*規格", "需確認\\庫存")
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateSpecialCharSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSpecialCharSheet = ws
End Function