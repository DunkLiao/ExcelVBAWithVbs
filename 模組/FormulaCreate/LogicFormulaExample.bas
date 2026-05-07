Attribute VB_Name = "LogicFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: LogicFormulaExample
'功能描述: 在 Excel 中示範邏輯判斷公式的使用範例
'          包含 IF、AND、OR、NOT、IFS、SWITCH 等公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

Sub TestLogicFormula()
    Call CreateLogicFormulaExample("邏輯公式範例")
End Sub

Sub CreateLogicFormulaExample(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "學生"
    ws.Range("B1").Value = "分數"
    ws.Range("C1").Value = "IF 及格判斷"
    ws.Range("D1").Value = "IFS 等第"
    ws.Range("E1").Value = "AND 條件"
    ws.Range("F1").Value = "OR 條件"
    ws.Range("A1:F1").Font.Bold = True
    Dim names(4) As String
    Dim scores(4) As Integer
    names(0) = "張三" : scores(0) = 92
    names(1) = "李四" : scores(1) = 58
    names(2) = "王五" : scores(2) = 75
    names(3) = "趙六" : scores(3) = 45
    names(4) = "陳七" : scores(4) = 88
    Dim i As Integer
    For i = 0 To 4
        Dim r As Integer
        r = i + 2
        ws.Cells(r, 1).Value = names(i)
        ws.Cells(r, 2).Value = scores(i)
        ws.Cells(r, 3).Formula = "=IF(B" & r & ">=60,""及格"",""不及格"")"
        ws.Cells(r, 3).Interior.Color = RGB(198, 239, 206)
        ws.Cells(r, 4).Formula = "=IFS(B" & r & ">=90,""A"",B" & r & ">=80,""B"",B" & r & ">=70,""C"",B" & r & ">=60,""D"",TRUE,""F"")"
        ws.Cells(r, 4).Interior.Color = RGB(255, 235, 156)
        ws.Cells(r, 5).Formula = "=AND(B" & r & ">=60,B" & r & "<=90)"
        ws.Cells(r, 5).Interior.Color = RGB(255, 199, 206)
        ws.Cells(r, 6).Formula = "=OR(B" & r & "<60,B" & r & ">=90)"
        ws.Cells(r, 6).Interior.Color = RGB(255, 199, 206)
    Next i
    ws.Range("A9").Value = "NOT 示範:"
    ws.Range("B9").Value = True
    ws.Range("C9").Formula = "=NOT(B9)"
    ws.Range("C9").Interior.Color = RGB(198, 239, 206)
    ws.Range("A11").Value = "SWITCH 示範:"
    ws.Range("B11").Value = 2
    ws.Range("C11").Formula = "=SWITCH(B11,1,""一月"",2,""二月"",3,""三月"",""其他"")"
    ws.Range("C11").Interior.Color = RGB(198, 239, 206)
    ws.Columns("A:F").AutoFit
    MsgBox "邏輯公式範例已建立完成！", vbInformation, "完成"
End Sub
