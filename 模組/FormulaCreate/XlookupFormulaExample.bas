Option Explicit
Attribute VB_Name = "XlookupFormulaExample"
'*************************************************************************************
'家舱嘿: XLOOKUP そΑ絛ㄒ
'弧:  VBA уΩ糶 XLOOKUP そΑボ絛弘絋琩高籔家絢ゑ癸
'
'舦┮Τ: Dunk
'祘Α砞璸: Dunk
'级糶ら戳: 2026/5/27
'
'*************************************************************************************
Sub TestXlookupFormula()
    Call CreateXlookupFormulaExample("XLOOKUPそΑ絛ㄒ")
End Sub

Sub CreateXlookupFormulaExample(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheetXlookup(sheetName)
    ws.Cells.Clear

    Call FillXlookupData(ws)
    Call WriteXlookupFormulas(ws)

    ws.Columns("A:G").AutoFit
    MsgBox "XLOOKUP そΑ絛ㄒミЧΘ", vbInformation, "ЧΘ"
    Exit Sub

ErrorHandler:
    MsgBox "ミ XLOOKUP そΑ祇ネ岿粇" & Err.Description, vbExclamation, "岿粇"
End Sub

Private Function GetOrCreateWorksheetXlookup(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheetXlookup = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheetXlookup Is Nothing Then
        Set GetOrCreateWorksheetXlookup = ThisWorkbook.Worksheets.Add
        GetOrCreateWorksheetXlookup.Name = sheetName
    End If
End Function

Private Sub FillXlookupData(ByVal ws As Worksheet)
    ' ㄓ方戈 A1:C6
    ws.Range("A1").Value = "絪腹"
    ws.Range("B1").Value = "﹎"
    ws.Range("C1").Value = "场"

    ws.Range("A2").Value = "E001"
    ws.Range("B2").Value = ""
    ws.Range("C2").Value = "穨叭场"

    ws.Range("A3").Value = "E002"
    ws.Range("B3").Value = "地"
    ws.Range("C3").Value = "癩叭场"

    ws.Range("A4").Value = "E003"
    ws.Range("B4").Value = "眎"
    ws.Range("C4").Value = "戈场"

    ws.Range("A5").Value = "E004"
    ws.Range("B5").Value = "朝Щ"
    ws.Range("C5").Value = "祇场"

    ws.Range("A6").Value = "E005"
    ws.Range("B6").Value = "狶睶"
    ws.Range("C6").Value = "︽綪场"

    ' 琩高跋 E1:G3
    ws.Range("E1").Value = "琩高絪腹"
    ws.Range("F1").Value = "琩高﹎"
    ws.Range("G1").Value = "琩高场"

    ws.Range("E2").Value = "E003"
    ws.Range("E3").Value = "E005"
End Sub

Private Sub WriteXlookupFormulas(ByVal ws As Worksheet)
    ' 弘絋琩高﹎
    ws.Range("F2").Formula = "=XLOOKUP(E2,A2:A6,B2:B6,""тぃ"",0)"
    ws.Range("F3").Formula = "=XLOOKUP(E3,A2:A6,B2:B6,""тぃ"",0)"

    ' 弘絋琩高场
    ws.Range("G2").Formula = "=XLOOKUP(E2,A2:A6,C2:C6,""тぃ"",0)"
    ws.Range("G3").Formula = "=XLOOKUP(E3,A2:A6,C2:C6,""тぃ"",0)"

    ' 弧
    ws.Range("E5").Value = "弧XLOOKUP(琩高, 琩高絛瞅, 肚絛瞅, тぃ, ゑ癸家Α)"
End Sub
