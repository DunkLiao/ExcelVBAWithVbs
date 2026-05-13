Attribute VB_Name = "InformationFormulaExample"
Option Explicit
'*************************************************************************************
'家舱嘿: InformationFormulaExample
'弧:  Excel 纗いミ戈癟ㄧ计そΑ絛ㄒ
'           ISBLANKISNUMBERISTEXTISERRORISODDISEVEN 单ㄧ计
'
'舦┮Τ: Dunk
'祘Α砞璸: Dunk
'级糶ら戳: 2026/5/13
'
'*************************************************************************************

Sub TestInformationFormula()
    Call CreateInformationFormulaExample("戈癟ㄧ计絛ㄒ")
End Sub

Sub CreateInformationFormulaExample(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    Call FillInfoData(ws)

    ' ISBLANK 耞琌フ
    ws.Range("C2").Formula = "=ISBLANK(A2)"
    ' ISNUMBER 耞琌计
    ws.Range("C3").Formula = "=ISNUMBER(A3)"
    ' ISTEXT 耞琌ゅ
    ws.Range("C4").Formula = "=ISTEXT(A4)"
    ' ISERROR 耞琌岿粇
    ws.Range("C5").Formula = "=ISERROR(A5)"
    ' ISODD 耞琌计
    ws.Range("C6").Formula = "=ISODD(A6)"
    ' ISEVEN 耞琌案计
    ws.Range("C7").Formula = "=ISEVEN(A7)"

    ws.Range("D2").Value = "ISBLANKフ耞"
    ws.Range("D3").Value = "ISNUMBER计耞"
    ws.Range("D4").Value = "ISTEXTゅ耞"
    ws.Range("D5").Value = "ISERROR岿粇耞"
    ws.Range("D6").Value = "ISODD计耞"
    ws.Range("D7").Value = "ISEVEN案计耞"

    ws.Range("C2:C7").Interior.Color = RGB(198, 239, 206)
    ws.Columns("A:D").AutoFit
    MsgBox "戈癟ㄧ计絛ㄒミЧΘ", vbInformation, "ЧΘ"
End Sub

Private Sub FillInfoData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "代刚"
    ws.Range("B1").Value = "弧"
    ws.Range("C1").Value = "挡狦"
    ws.Range("D1").Value = "ㄧ计"

    ws.Range("A2").Value = ""
    ws.Range("B2").Value = "フ纗"
    ws.Range("A3").Value = 123
    ws.Range("B3").Value = "计 123"
    ws.Range("A4").Value = "Hello"
    ws.Range("B4").Value = "ゅ Hello"
    ws.Range("A5").Formula = "=1/0"
    ws.Range("B5").Value = "埃箂岿粇"
    ws.Range("A6").Value = 7
    ws.Range("B6").Value = "计 7"
    ws.Range("A7").Value = 8
    ws.Range("B7").Value = "案计 8"

    ws.Range("A1:D1").Font.Bold = True
End Sub
