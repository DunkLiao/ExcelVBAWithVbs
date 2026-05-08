Option Explicit

' 建立常見儀表板公式範例，包含 SUMIFS、COUNTIFS、AVERAGEIFS 與 IFERROR。
Public Sub CreateDashboardFormulaExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim sheetName As String

    sheetName = "儀表板公式範例"
    Set ws = GetOrCreateFormulaWorksheet(sheetName)
    ws.Cells.Clear

    Call FillFormulaSourceData(ws)
    Call FillFormulaSummaryArea(ws)
    ws.Columns("A:H").AutoFit

    MsgBox "儀表板公式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立公式範例失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillFormulaSourceData(ByVal ws As Worksheet)
    ws.Range("A1:E1").Value = Array("日期", "區域", "產品", "數量", "金額")
    ws.Range("A2:E2").Value = Array(DateSerial(2026, 1, 3), "北區", "A", 12, 3600)
    ws.Range("A3:E3").Value = Array(DateSerial(2026, 1, 6), "中區", "B", 8, 2800)
    ws.Range("A4:E4").Value = Array(DateSerial(2026, 1, 10), "北區", "B", 10, 3500)
    ws.Range("A5:E5").Value = Array(DateSerial(2026, 1, 18), "南區", "A", 15, 4500)
    ws.Range("A6:E6").Value = Array(DateSerial(2026, 2, 2), "北區", "A", 9, 2700)
    ws.Range("A7:E7").Value = Array(DateSerial(2026, 2, 9), "中區", "A", 11, 3300)
    ws.Range("A8:E8").Value = Array(DateSerial(2026, 2, 14), "南區", "B", 7, 2450)
    ws.Range("A2:A8").NumberFormat = "yyyy/mm/dd"
End Sub

Private Sub FillFormulaSummaryArea(ByVal ws As Worksheet)
    ws.Range("G1:H1").Value = Array("指標", "公式結果")
    ws.Range("G2").Value = "北區總金額"
    ws.Range("H2").Formula = "=SUMIFS(E:E,B:B,""北區"")"
    ws.Range("G3").Value = "產品A筆數"
    ws.Range("H3").Formula = "=COUNTIFS(C:C,""A"")"
    ws.Range("G4").Value = "中區平均金額"
    ws.Range("H4").Formula = "=IFERROR(AVERAGEIFS(E:E,B:B,""中區""),0)"
    ws.Range("G5").Value = "一月總金額"
    ws.Range("H5").Formula = "=SUMIFS(E:E,A:A,"">=""&DATE(2026,1,1),A:A,""<""&DATE(2026,2,1))"
    ws.Range("H2:H5").NumberFormat = "#,##0"
End Sub

Private Function GetOrCreateFormulaWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateFormulaWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateFormulaWorksheet Is Nothing Then
        Set GetOrCreateFormulaWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateFormulaWorksheet.Name = sheetName
    End If
End Function