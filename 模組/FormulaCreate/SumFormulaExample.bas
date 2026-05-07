Attribute VB_Name = "SumFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: SumFormulaExample
'功能描述: 在 Excel 中示範加總相關公式的使用範例
'          包含 SUM、SUMIF、SUMIFS 等公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

Sub TestSumFormula()
    Call CreateSumFormulaExample("加總公式範例")
End Sub

Sub CreateSumFormulaExample(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Call FillSumData(ws)
    ws.Range("D2").Formula = "=SUM(C2:C11)"
    ws.Range("D2").Interior.Color = RGB(198, 239, 206)
    ws.Range("D4").Formula = "=SUMIF(B2:B11,""業務"",C2:C11)"
    ws.Range("D4").Interior.Color = RGB(198, 239, 206)
    ws.Range("D6").Formula = "=SUMIFS(C2:C11,B2:B11,""業務"",C2:C11,"">500"")"
    ws.Range("D6").Interior.Color = RGB(198, 239, 206)
    ws.Range("E2").Value = "SUM 全部合計"
    ws.Range("E4").Value = "SUMIF 業務部門合計"
    ws.Range("E6").Value = "SUMIFS 業務部門且>500合計"
    ws.Columns("A:E").AutoFit
    MsgBox "加總公式範例已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillSumData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "銷售金額"
    Dim data(9, 2) As Variant
    data(0, 0) = "張三" : data(0, 1) = "業務" : data(0, 2) = 800
    data(1, 0) = "李四" : data(1, 1) = "行政" : data(1, 2) = 300
    data(2, 0) = "王五" : data(2, 1) = "業務" : data(2, 2) = 650
    data(3, 0) = "趙六" : data(3, 1) = "研發" : data(3, 2) = 450
    data(4, 0) = "陳七" : data(4, 1) = "業務" : data(4, 2) = 920
    data(5, 0) = "林八" : data(5, 1) = "行政" : data(5, 2) = 280
    data(6, 0) = "吳九" : data(6, 1) = "研發" : data(6, 2) = 510
    data(7, 0) = "黃十" : data(7, 1) = "業務" : data(7, 2) = 380
    data(8, 0) = "周一" : data(8, 1) = "業務" : data(8, 2) = 740
    data(9, 0) = "馮二" : data(9, 1) = "研發" : data(9, 2) = 600
    Dim i As Integer
    For i = 0 To 9
        ws.Cells(i + 2, 1).Value = data(i, 0)
        ws.Cells(i + 2, 2).Value = data(i, 1)
        ws.Cells(i + 2, 3).Value = data(i, 2)
    Next i
    ws.Range("A1:C1").Font.Bold = True
End Sub
