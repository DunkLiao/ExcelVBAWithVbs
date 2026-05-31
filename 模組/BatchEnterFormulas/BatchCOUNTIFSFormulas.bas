Attribute VB_Name = "BatchCOUNTIFSFormulas"
Option Explicit
'*************************************************************************************
'家舱嘿: BatchCOUNTIFSFormulas
'弧: уΩ块COUNTIFS兵ン璸计そΑ参璸才兵ン癘魁计秖
'
'舦┮Τ: Dunk
'祘Α砞璸: Dunk
'级糶ら戳: 2026/5/31
'
'*************************************************************************************

Sub TestBatchCOUNTIFSFormulas()
    Dim ws As Worksheet
    Set ws = GetOrCreateCountIfsSheet(ThisWorkbook, "COUNTIFSそΑ絛ㄒ")
    Call FillCOUNTIFSSampleData(ws)
    Call InsertBatchCOUNTIFSFormulas(ws)
    MsgBox "COUNTIFSそΑуΩミЧΘ", vbInformation, "ЧΘ"
End Sub

Sub InsertBatchCOUNTIFSFormulas(ByVal ws As Worksheet)
    ws.Range("F1").Value = "参璸兵ン"
    ws.Range("G1").Value = "计秖"

    ws.Range("F2").Value = "穨叭场-ЧΘ"
    ws.Range("G2").Formula = "=COUNTIFS(A:A,""穨叭场"",D:D,""ЧΘ"")"

    ws.Range("F3").Value = "︽綪场-ЧΘ"
    ws.Range("G3").Formula = "=COUNTIFS(A:A,""︽綪场"",D:D,""ЧΘ"")"

    ws.Range("F4").Value = "穨叭场-秈︽い"
    ws.Range("G4").Formula = "=COUNTIFS(A:A,""穨叭场"",D:D,""秈︽い"")"

    ws.Range("F5").Value = "肂>50000-ЧΘ"
    ws.Range("G5").Formula = "=COUNTIFS(C:C,"">50000"",D:D,""ЧΘ"")"

    ws.Range("F6").Value = "る-穨叭场"
    ws.Range("G6").Formula = "=COUNTIFS(B:B,""る"",A:A,""穨叭场"")"

    ws.Range("F7").Value = "肂30000-80000"
    ws.Range("G7").Formula = "=COUNTIFS(C:C,"">=30000"",C:C,""<=80000"")"

    ws.Columns("F:G").AutoFit
End Sub

Private Sub FillCOUNTIFSSampleData(ByVal ws As Worksheet)
    Dim dataArr As Variant
    Dim i       As Integer

    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("场", "る", "肂", "篈")

    dataArr = Array( _
        Array("穨叭场", "る", 45000, "ЧΘ"), _
        Array("︽綪场", "る", 62000, "ЧΘ"), _
        Array("穨叭场", "る", 38000, "秈︽い"), _
        Array("穨叭场", "る", 75000, "ЧΘ"), _
        Array("︽綪场", "る", 28000, ""), _
        Array("穨叭场", "る", 55000, "ЧΘ"), _
        Array("︽綪场", "る", 41000, "秈︽い"), _
        Array("穨叭场", "る", 83000, "ЧΘ"))

    For i = 0 To UBound(dataArr)
        ws.Cells(i + 2, 1).Value = dataArr(i)(0)
        ws.Cells(i + 2, 2).Value = dataArr(i)(1)
        ws.Cells(i + 2, 3).Value = dataArr(i)(2)
        ws.Cells(i + 2, 4).Value = dataArr(i)(3)
    Next i

    ws.Columns("A:D").AutoFit
End Sub

Private Function GetOrCreateCountIfsSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateCountIfsSheet = ws
End Function
