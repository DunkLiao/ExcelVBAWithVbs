Attribute VB_Name = "LookupFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: LookupFormulaExample
'功能描述: 在 Excel 中示範查閱參照公式的使用範例
'          包含 VLOOKUP、HLOOKUP、INDEX、MATCH 等公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

Sub TestLookupFormula()
    Call CreateLookupFormulaExample("查閱公式範例")
End Sub

Sub CreateLookupFormulaExample(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Call FillLookupData(ws)
    ws.Range("F2").Formula = "=VLOOKUP(E2,A2:C11,2,FALSE)"
    ws.Range("F2").Interior.Color = RGB(198, 239, 206)
    ws.Range("G2").Formula = "=VLOOKUP(E2,A2:C11,3,FALSE)"
    ws.Range("G2").Interior.Color = RGB(198, 239, 206)
    ws.Range("F4").Formula = "=INDEX(C2:C11,MATCH(E4,B2:B11,0))"
    ws.Range("F4").Interior.Color = RGB(255, 235, 156)
    ws.Range("E1").Value = "查詢條件"
    ws.Range("E2").Value = "E003"
    ws.Range("E4").Value = "王五"
    ws.Range("F1").Value = "VLOOKUP 姓名"
    ws.Range("G1").Value = "VLOOKUP 薪資"
    ws.Range("F3").Value = "INDEX+MATCH 薪資"
    ws.Range("E1:G1").Font.Bold = True
    ws.Columns("A:G").AutoFit
    MsgBox "查閱公式範例已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillLookupData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "員工編號"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "薪資"
    Dim data(9, 2) As Variant
    data(0, 0) = "E001" : data(0, 1) = "張三" : data(0, 2) = 45000
    data(1, 0) = "E002" : data(1, 1) = "李四" : data(1, 2) = 38000
    data(2, 0) = "E003" : data(2, 1) = "王五" : data(2, 2) = 52000
    data(3, 0) = "E004" : data(3, 1) = "趙六" : data(3, 2) = 41000
    data(4, 0) = "E005" : data(4, 1) = "陳七" : data(4, 2) = 67000
    data(5, 0) = "E006" : data(5, 1) = "林八" : data(5, 2) = 35000
    data(6, 0) = "E007" : data(6, 1) = "吳九" : data(6, 2) = 48000
    data(7, 0) = "E008" : data(7, 1) = "黃十" : data(7, 2) = 55000
    data(8, 0) = "E009" : data(8, 1) = "周一" : data(8, 2) = 43000
    data(9, 0) = "E010" : data(9, 1) = "馮二" : data(9, 2) = 60000
    Dim i As Integer
    For i = 0 To 9
        ws.Cells(i + 2, 1).Value = data(i, 0)
        ws.Cells(i + 2, 2).Value = data(i, 1)
        ws.Cells(i + 2, 3).Value = data(i, 2)
    Next i
    ws.Range("A1:C1").Font.Bold = True
End Sub
