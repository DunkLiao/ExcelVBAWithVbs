Option Explicit

' 批次輸入查找公式，依產品代碼帶出產品名稱與單價。
Public Sub BatchEnterLookupFormulasExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = GetOrCreateLookupWorksheet("批次查找公式範例")
    ws.Cells.Clear
    Call FillLookupFormulaData(ws)

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("B2:B" & lastRow).Formula = "=IFERROR(VLOOKUP(A2,$F$2:$H$5,2,FALSE),""查無資料"")"
    ws.Range("C2:C" & lastRow).Formula = "=IFERROR(VLOOKUP(A2,$F$2:$H$5,3,FALSE),0)"
    ws.Range("D2:D" & lastRow).Formula = "=C2*E2"
    ws.Range("C2:D" & lastRow).NumberFormat = "#,##0"
    ws.Columns("A:H").AutoFit

    MsgBox "查找公式已批次輸入完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "批次輸入公式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillLookupFormulaData(ByVal ws As Worksheet)
    ws.Range("A1:E1").Value = Array("產品代碼", "產品名稱", "單價", "小計", "數量")
    ws.Range("A2:E2").Value = Array("P001", "", "", "", 3)
    ws.Range("A3:E3").Value = Array("P003", "", "", "", 5)
    ws.Range("A4:E4").Value = Array("P002", "", "", "", 2)
    ws.Range("A5:E5").Value = Array("P004", "", "", "", 4)
    ws.Range("F1:H1").Value = Array("產品代碼", "產品名稱", "單價")
    ws.Range("F2:H2").Value = Array("P001", "鍵盤", 900)
    ws.Range("F3:H3").Value = Array("P002", "滑鼠", 450)
    ws.Range("F4:H4").Value = Array("P003", "螢幕", 5200)
    ws.Range("F5:H5").Value = Array("P004", "耳機", 1300)
End Sub

Private Function GetOrCreateLookupWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateLookupWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateLookupWorksheet Is Nothing Then
        Set GetOrCreateLookupWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateLookupWorksheet.Name = sheetName
    End If
End Function