Attribute VB_Name = "BatchFrequencyFormulas"
Option Explicit
'*************************************************************************************
'家舱嘿: BatchFrequencyFormulas
'弧: 玻ネ妓セ计籔禯уΩ块 FREQUENCY 皚そΑ
'
'舦┮Τ: Dunk
'祘Α砞璸: Dunk
'级糶ら戳: 2026/5/15
'
'*************************************************************************************

Public Sub RunBatchFrequencyFormulas()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet

    Set ws = GetOrCreateFrequencySheet("Ω计だそΑ")
    ws.Cells.Clear

    Call FillFrequencySampleData(ws)
    Call FillFrequencyBins(ws)
    Call ApplyFrequencyFormula(ws)
    Call FillFrequencySummary(ws)

    ws.Columns("A:G").AutoFit
    MsgBox "FREQUENCY 皚そΑミЧΘ", vbInformation, "ЧΘ"
    Exit Sub

ErrorHandler:
    MsgBox "ミΩ计だそΑ祇ネ岿粇: " & Err.Description, vbExclamation, "岿粇"
End Sub

Private Sub FillFrequencySampleData(ByVal ws As Worksheet)
    Dim valuesData As Variant
    Dim i As Long

    valuesData = Array(12, 8, 15, 23, 28, 31, 35, 42, 18, 27, 30, 16, 9, 14, 22, _
                       26, 33, 37, 41, 45, 11, 19, 24, 29, 32, 36, 39, 44, 48, 52)

    ws.Range("A1").Value = "妓セ计"
    For i = LBound(valuesData) To UBound(valuesData)
        ws.Cells(i + 2, 1).Value = valuesData(i)
    Next i
End Sub

Private Sub FillFrequencyBins(ByVal ws As Worksheet)
    ws.Range("C1:D1").Value = Array("禯", "Ω计")
    ws.Range("C2").Value = 10
    ws.Range("C3").Value = 20
    ws.Range("C4").Value = 30
    ws.Range("C5").Value = 40
    ws.Range("C6").Value = 50
    ws.Range("B2").Value = "<=10"
    ws.Range("B3").Value = "11-20"
    ws.Range("B4").Value = "21-30"
    ws.Range("B5").Value = "31-40"
    ws.Range("B6").Value = "41-50"
    ws.Range("B7").Value = ">50"
End Sub

Private Sub ApplyFrequencyFormula(ByVal ws As Worksheet)
    ws.Range("D2:D7").ClearContents
    ws.Range("D2:D7").FormulaArray = "=FREQUENCY(A2:A31,C2:C6)"
End Sub

Private Sub FillFrequencySummary(ByVal ws As Worksheet)
    ws.Range("F1:G1").Value = Array("篕璶", "")
    ws.Range("F2").Value = "妓セ计秖"
    ws.Range("G2").Formula = "=COUNT(A2:A31)"
    ws.Range("F3").Value = "程"
    ws.Range("G3").Formula = "=MIN(A2:A31)"
    ws.Range("F4").Value = "程"
    ws.Range("G4").Formula = "=MAX(A2:A31)"
    ws.Range("F5").Value = "Ω计羆㎝"
    ws.Range("G5").Formula = "=SUM(D2:D7)"
End Sub

Private Function GetOrCreateFrequencySheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateFrequencySheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateFrequencySheet Is Nothing Then
        Set GetOrCreateFrequencySheet = ThisWorkbook.Worksheets.Add
        GetOrCreateFrequencySheet.Name = sheetName
    End If
End Function
