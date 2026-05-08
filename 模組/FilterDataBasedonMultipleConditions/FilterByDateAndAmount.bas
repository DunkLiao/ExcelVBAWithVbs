Option Explicit

' 依日期區間與最低金額篩選資料。
Public Sub FilterByDateAndAmountExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim minAmount As Double

    Set ws = GetOrCreateFilterWorksheet("日期金額篩選範例")
    ws.Cells.Clear
    Call FillDateAmountData(ws)

    startDate = DateSerial(2026, 1, 1)
    endDate = DateSerial(2026, 1, 31)
    minAmount = 3000
    Call ApplyDateAmountFilter(ws, startDate, endDate, minAmount)

    MsgBox "多重條件篩選已套用完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "篩選資料失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub ApplyDateAmountFilter(ByVal ws As Worksheet, ByVal startDate As Date, ByVal endDate As Date, ByVal minAmount As Double)
    Dim dataRange As Range

    Set dataRange = ws.Range("A1").CurrentRegion
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    dataRange.AutoFilter Field:=1, Criteria1:=">=" & CLng(startDate), Operator:=xlAnd, Criteria2:="<=" & CLng(endDate)
    dataRange.AutoFilter Field:=4, Criteria1:=">=" & minAmount
    ws.Columns("A:D").AutoFit
End Sub

Private Sub FillDateAmountData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("日期", "客戶", "產品", "金額")
    ws.Range("A2:D2").Value = Array(DateSerial(2026, 1, 3), "甲公司", "A", 2800)
    ws.Range("A3:D3").Value = Array(DateSerial(2026, 1, 8), "乙公司", "B", 5200)
    ws.Range("A4:D4").Value = Array(DateSerial(2026, 1, 20), "丙公司", "A", 3600)
    ws.Range("A5:D5").Value = Array(DateSerial(2026, 2, 5), "丁公司", "C", 4100)
    ws.Range("A6:D6").Value = Array(DateSerial(2026, 1, 26), "戊公司", "B", 1900)
    ws.Range("A2:A6").NumberFormat = "yyyy/mm/dd"
    ws.Range("D2:D6").NumberFormat = "#,##0"
End Sub

Private Function GetOrCreateFilterWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateFilterWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateFilterWorksheet Is Nothing Then
        Set GetOrCreateFilterWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateFilterWorksheet.Name = sheetName
    End If
End Function