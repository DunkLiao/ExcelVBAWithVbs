Attribute VB_Name = "TextSplitFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: TextSplitFormulaExample
'功能說明: 示範透過VBA輸入文字分割公式（TEXTSPLIT、TEXTBEFORE、TEXTAFTER）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestTextSplitFormulas()
    Call CreateTextSplitExample
End Sub

Sub CreateTextSplitExample()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim i As Long
    
    sheetName = "文字分割公式"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillTextSplitData(ws)
    
    ws.Range("C1").Value = "姓氏"
    ws.Range("D1").Value = "名字"
    ws.Range("E1").Value = "分割陣列"
    
    For i = 2 To 7
        ws.Cells(i, 3).Formula = "=TEXTBEFORE(A" & i & ","" "")"
        ws.Cells(i, 4).Formula = "=TEXTAFTER(A" & i & ","" "")"
    Next i
    
    ws.Range("E2").Formula = "=TEXTSPLIT(A2,"" "")"
    
    For i = 2 To 6
        ws.Cells(i + 6, 5).Formula = "=TEXTSPLIT(A" & i & ","" "")"
    Next i
    
    ws.Columns("A:E").AutoFit
    ws.Activate
    
    MsgBox "文字分割公式範例已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillTextSplitData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "原始姓名"
    ws.Range("B1").Value = "備註"
    
    ws.Range("A2").Value = "王 小明"
    ws.Range("B2").Value = "範例"
    ws.Range("A3").Value = "李 大華"
    ws.Range("B3").Value = "範例"
    ws.Range("A4").Value = "張 美麗"
    ws.Range("B4").Value = "範例"
    ws.Range("A5").Value = "陳 建國"
    ws.Range("B5").Value = "範例"
    ws.Range("A6").Value = "林 志玲"
    ws.Range("B6").Value = "範例"
    ws.Range("A7").Value = "周 杰倫"
    ws.Range("B7").Value = "範例"
    
    ws.Columns("A:B").AutoFit
End Sub
