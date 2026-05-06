Attribute VB_Name = "ErrorHandleFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: ErrorHandleFormulaExample
'功能描述: 在 Excel 中示範錯誤處理公式的使用範例
'          包含 IFERROR、ISERROR、ISBLANK、ISNA、ISNUMBER、ISTEXT 等公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

Sub TestErrorHandleFormula()
    Call CreateErrorHandleFormulaExample("錯誤處理公式範例")
End Sub

Sub CreateErrorHandleFormulaExample(ByVal sheetName As String)
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
    ws.Range("B1").Value = "查閱對象"
    ws.Range("C1").Value = "IFERROR 保護"
    ws.Range("D1").Value = "ISERROR 偵測"
    ws.Range("E1").Value = "ISBLANK"
    ws.Range("F1").Value = "ISNUMBER"
    ws.Range("G1").Value = "ISTEXT"
    ws.Range("A1:G1").Font.Bold = True
    ws.Range("H1").Value = "查閱表"
    ws.Range("H2").Value = "A001"
    ws.Range("H3").Value = "A002"
    ws.Range("H4").Value = "A003"
    ws.Range("I1").Value = "名稱"
    ws.Range("I2").Value = "蘋果"
    ws.Range("I3").Value = "橘子"
    ws.Range("I4").Value = "葡萄"
    ws.Range("A2").Value = 123
    ws.Range("A3").Value = "文字"
    ws.Range("A4").Value = ""
    ws.Range("A5").Value = 0
    ws.Range("B2").Value = "A001"
    ws.Range("B3").Value = "A999"
    ws.Range("B4").Value = "A002"
    ws.Range("B5").Value = "A003"
    Dim i As Integer
    For i = 2 To 5
        ws.Cells(i, 3).Formula = "=IFERROR(VLOOKUP(B" & i & ",H$2:I$4,2,FALSE),""查無資料"")"
        ws.Cells(i, 3).Interior.Color = RGB(198, 239, 206)
        ws.Cells(i, 4).Formula = "=ISERROR(VLOOKUP(B" & i & ",H$2:I$4,2,FALSE))"
        ws.Cells(i, 4).Interior.Color = RGB(255, 235, 156)
        ws.Cells(i, 5).Formula = "=ISBLANK(A" & i & ")"
        ws.Cells(i, 5).Interior.Color = RGB(255, 199, 206)
        ws.Cells(i, 6).Formula = "=ISNUMBER(A" & i & ")"
        ws.Cells(i, 6).Interior.Color = RGB(255, 199, 206)
        ws.Cells(i, 7).Formula = "=ISTEXT(A" & i & ")"
        ws.Cells(i, 7).Interior.Color = RGB(255, 199, 206)
    Next i
    ws.Range("A7").Value = "ISNA 示範:"
    ws.Range("B7").Value = "A999"
    ws.Range("C7").Formula = "=ISNA(VLOOKUP(B7,H$2:I$4,2,FALSE))"
    ws.Range("C7").Interior.Color = RGB(198, 239, 206)
    ws.Columns("A:I").AutoFit
    MsgBox "錯誤處理公式範例已建立完成！", vbInformation, "完成"
End Sub
