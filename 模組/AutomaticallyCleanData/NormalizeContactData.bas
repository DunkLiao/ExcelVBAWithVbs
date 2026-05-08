Option Explicit

' 清理電話與電子郵件欄位，移除多餘空白並統一大小寫。
Public Sub NormalizeContactDataExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet

    Set ws = GetOrCreateCleanWorksheet("聯絡資料清理範例")
    ws.Cells.Clear
    Call FillContactData(ws)
    Call NormalizeContactData(ws)

    ws.Columns("A:D").AutoFit
    MsgBox "聯絡資料清理完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清理聯絡資料失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub NormalizeContactData(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim phoneText As String
    Dim emailText As String

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For rowIndex = 2 To lastRow
        phoneText = CStr(ws.Cells(rowIndex, "B").Value)
        phoneText = Replace(phoneText, " ", "")
        phoneText = Replace(phoneText, "-", "")
        ws.Cells(rowIndex, "C").Value = phoneText

        emailText = Trim$(CStr(ws.Cells(rowIndex, "D").Value))
        ws.Cells(rowIndex, "D").Value = LCase$(emailText)
    Next rowIndex
End Sub

Private Sub FillContactData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("姓名", "原始電話", "清理後電話", "電子郵件")
    ws.Range("A2:D2").Value = Array("王小明", "0912-345-678", "", " USER1@EXAMPLE.COM ")
    ws.Range("A3:D3").Value = Array("陳美華", " 02-2345-6789", "", "Sales@Example.com")
    ws.Range("A4:D4").Value = Array("林志強", "07 2233 456", "", " SERVICE@EXAMPLE.COM")
End Sub

Private Function GetOrCreateCleanWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateCleanWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateCleanWorksheet Is Nothing Then
        Set GetOrCreateCleanWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateCleanWorksheet.Name = sheetName
    End If
End Function