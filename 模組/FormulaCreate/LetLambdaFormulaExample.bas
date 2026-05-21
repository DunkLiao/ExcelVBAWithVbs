Option Explicit
Attribute VB_Name = "LetLambdaFormulaExample"
'*************************************************************************************
'模組名稱: LetLambdaFormulaExample
'功能說明: 示範透過 VBA 批次插入 Excel LET 函數公式，計算折扣金額與分類
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestLetLambdaFormula()
    Call CreateLetFormulaSheet("LET公式範例")
End Sub

Sub CreateLetFormulaSheet(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateLetSheet(sheetName)
    ws.Cells.Clear

    Call FillLetSampleData(ws)
    Call InsertLetFormulas(ws)

    ws.Columns("A:F").AutoFit
    MsgBox "LET 函數公式已插入完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "插入 LET 公式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillLetSampleData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "品項"
    ws.Range("B1").Value = "單價"
    ws.Range("C1").Value = "數量"
    ws.Range("D1").Value = "折扣率"
    ws.Range("E1").Value = "LET計算金額"
    ws.Range("F1").Value = "LET金額分類"

    ws.Range("A2").Value = "筆記型電腦"
    ws.Range("B2").Value = 35000
    ws.Range("C2").Value = 2
    ws.Range("D2").Value = 0.9

    ws.Range("A3").Value = "平板電腦"
    ws.Range("B3").Value = 18000
    ws.Range("C3").Value = 5
    ws.Range("D3").Value = 0.85

    ws.Range("A4").Value = "智慧型手機"
    ws.Range("B4").Value = 28000
    ws.Range("C4").Value = 3
    ws.Range("D4").Value = 0.88

    ws.Range("A5").Value = "無線耳機"
    ws.Range("B5").Value = 3500
    ws.Range("C5").Value = 10
    ws.Range("D5").Value = 0.95

    ws.Range("A6").Value = "滑鼠"
    ws.Range("B6").Value = 800
    ws.Range("C6").Value = 20
    ws.Range("D6").Value = 1
End Sub

Private Sub InsertLetFormulas(ByVal ws As Worksheet)
    Dim r As Integer
    For r = 2 To 6
        ws.Cells(r, 5).Formula = _
            "=LET(p,B" & r & ",q,C" & r & ",d,D" & r & ",p*q*d)"
        ws.Cells(r, 6).Formula = _
            "=LET(a,E" & r & _
            ",IF(a>=50000," & Chr(34) & "高額" & Chr(34) & _
            ",IF(a>=10000," & Chr(34) & "中額" & Chr(34) & _
            "," & Chr(34) & "一般" & Chr(34) & ")))"
    Next r
End Sub

Private Function GetOrCreateLetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateLetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateLetSheet Is Nothing Then
        Set GetOrCreateLetSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateLetSheet.Name = sheetName
    End If
End Function
