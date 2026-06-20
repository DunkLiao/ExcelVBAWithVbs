Attribute VB_Name = "ForecastFormulaExample"
Option Explicit
'*************************************************************************************
'家舱嘿: ForecastFormulaExample
'弧: ボ絛硓筁VBAExcelい块箇代そΑFORECAST.LINEARTRENDGROWTH絛ㄒ祘Α
'
'舦┮Τ: Dunk
'祘Α砞璸: Dunk
'级糶ら戳: 2026/6/19
'
'*************************************************************************************

' 虏て代刚
Sub TestForecastFormula()
    Call CreateForecastFormulaExample
End Sub

' ミ箇代そΑ絛ㄒ
Sub CreateForecastFormulaExample()
    Dim ws As Worksheet
    Dim sheetName As String
    
    sheetName = "箇代そΑ絛ㄒ"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillForecastData(ws)
    Call EnterForecastFormulas(ws)
    
    ws.Activate
    MsgBox "箇代そΑ絛ㄒミЧΘ", vbInformation, "ЧΘ"
End Sub

' 恶箇代菌戈
Private Sub FillForecastData(ByVal ws As Worksheet)
    Dim i As Long
    
    ws.Range("A1").Value = "る"
    ws.Range("B1").Value = "綪扳肂"
    
    ' ㄏノ癹伴恶戈
    For i = 1 To 12
        ws.Cells(i + 1, 1).Value = "2024/" & i
        ws.Cells(i + 1, 2).Value = 1000 + i * 100 + i * 10
    Next i
End Sub

' 块箇代そΑ
Private Sub EnterForecastFormulas(ByVal ws As Worksheet)
    ' 箇代ヘ夹跋
    ws.Range("D1").Value = "箇代ヘ夹る"
    ws.Range("E1").Value = "FORECAST.LINEAR"
    ws.Range("F1").Value = "TREND翴箇代"
    ws.Range("G1").Value = "GROWTH计箇代"
    
    ws.Range("D2").Value = "2025/1"
    ws.Range("D3").Value = "2025/2"
    ws.Range("D4").Value = "2025/3"
    
    ' FORECAST.LINEAR 絬┦箇代 (Excel 2016+)
    ws.Range("E2").Formula = "=FORECAST.LINEAR(D2,B2:B13,A2:A13)"
    ws.Range("E3").Formula = "=FORECAST.LINEAR(D3,B2:B13,A2:A13)"
    ws.Range("E4").Formula = "=FORECAST.LINEAR(D4,B2:B13,A2:A13)"
    
    ' TREND ㄧ计箇代翴
    ws.Range("F2").Formula = "=TREND(B2:B13,A2:A13,D2)"
    ws.Range("F3").Formula = "=TREND(B2:B13,A2:A13,D3)"
    ws.Range("F4").Formula = "=TREND(B2:B13,A2:A13,D4)"
    
    ' GROWTH 计Θ箇代
    ws.Range("G2").Formula = "=GROWTH(B2:B13,A2:A13,D2)"
    ws.Range("G3").Formula = "=GROWTH(B2:B13,A2:A13,D3)"
    ws.Range("G4").Formula = "=GROWTH(B2:B13,A2:A13,D4)"
    
    ws.Columns("A:G").AutoFit
End Sub
