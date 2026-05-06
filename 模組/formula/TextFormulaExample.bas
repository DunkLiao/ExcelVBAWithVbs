Attribute VB_Name = "TextFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: TextFormulaExample
'功能描述: 在 Excel 中示範文字處理公式的使用範例
'          包含 LEFT、RIGHT、MID、LEN、CONCATENATE、TRIM、TEXT 等公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

Sub TestTextFormula()
    Call CreateTextFormulaExample("文字公式範例")
End Sub

Sub CreateTextFormulaExample(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "原始資料"
    ws.Range("B1").Value = "LEFT(3)"
    ws.Range("C1").Value = "RIGHT(3)"
    ws.Range("D1").Value = "MID(2,4)"
    ws.Range("E1").Value = "LEN"
    ws.Range("F1").Value = "UPPER"
    ws.Range("G1").Value = "LOWER"
    ws.Range("A1:G1").Font.Bold = True
    Dim samples(4) As String
    samples(0) = "Hello World"
    samples(1) = "ExcelVBA"
    samples(2) = "  TRIM TEST  "
    samples(3) = "2026/05/07"
    samples(4) = "OpenAI GPT"
    Dim i As Integer
    For i = 0 To 4
        Dim r As Integer
        r = i + 2
        ws.Cells(r, 1).Value = samples(i)
        ws.Cells(r, 2).Formula = "=LEFT(A" & r & ",3)"
        ws.Cells(r, 3).Formula = "=RIGHT(A" & r & ",3)"
        ws.Cells(r, 4).Formula = "=MID(A" & r & ",2,4)"
        ws.Cells(r, 5).Formula = "=LEN(A" & r & ")"
        ws.Cells(r, 6).Formula = "=UPPER(A" & r & ")"
        ws.Cells(r, 7).Formula = "=LOWER(A" & r & ")"
    Next i
    ws.Range("A9").Value = "TRIM 示範:"
    ws.Range("B9").Value = "  多餘空格  "
    ws.Range("C9").Formula = "=TRIM(B9)"
    ws.Range("C9").Interior.Color = RGB(198, 239, 206)
    ws.Range("A11").Value = "CONCATENATE 示範:"
    ws.Range("B11").Value = "Excel"
    ws.Range("C11").Value = "VBA"
    ws.Range("D11").Formula = "=CONCATENATE(B11,""-"",C11)"
    ws.Range("D11").Interior.Color = RGB(198, 239, 206)
    ws.Range("A13").Value = "TEXT 示範:"
    ws.Range("B13").Value = 45388
    ws.Range("C13").Formula = "=TEXT(B13,""yyyy/mm/dd"")"
    ws.Range("C13").Interior.Color = RGB(198, 239, 206)
    ws.Columns("A:G").AutoFit
    MsgBox "文字公式範例已建立完成！", vbInformation, "完成"
End Sub
