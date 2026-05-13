Attribute VB_Name = "BatchWorkdayFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: 批次工作日公式
'功能說明: 批次在指定欄位插入WORKDAY、NETWORKDAYS等工作日計算公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub TestBatchWorkdayFormulas()
    Call CreateWorkdayFormulaDemo("工作日公式範例")
End Sub

Sub CreateWorkdayFormulaDemo(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    ws.Range("A1").Value = "開始日期"
    ws.Range("B1").Value = "工期(天)"
    ws.Range("C1").Value = "預計完成日(WORKDAY)"
    ws.Range("D1").Value = "結束日期"
    ws.Range("E1").Value = "實際工作天數(NETWORKDAYS)"
    ws.Range("F1").Value = "下個月第一個工作日"

    ws.Range("A2").Value = CDate("2026/5/1")  : ws.Range("B2").Value = 10 : ws.Range("D2").Value = CDate("2026/5/20")
    ws.Range("A3").Value = CDate("2026/6/1")  : ws.Range("B3").Value = 15 : ws.Range("D3").Value = CDate("2026/6/30")
    ws.Range("A4").Value = CDate("2026/7/1")  : ws.Range("B4").Value = 5  : ws.Range("D4").Value = CDate("2026/7/10")
    ws.Range("A5").Value = CDate("2026/8/1")  : ws.Range("B5").Value = 20 : ws.Range("D5").Value = CDate("2026/8/31")

    ws.Range("A2:A5").NumberFormat = "yyyy/m/d"
    ws.Range("D2:D5").NumberFormat = "yyyy/m/d"

    Call BatchInsertWorkdayFormula(ws, "C", 2, 5)
    Call BatchInsertNetworkdaysFormula(ws, "E", 2, 5)
    Call BatchInsertNextMonthFirstWorkdayFormula(ws, "F", 2, 5)

    ws.Columns("A:F").AutoFit
    ws.Range("A1:F1").Font.Bold = True

    MsgBox "工作日公式已批次建立完成！", vbInformation, "完成"
End Sub

Private Sub BatchInsertWorkdayFormula( _
    ByVal ws As Worksheet, _
    ByVal colLetter As String, _
    ByVal startRow As Integer, _
    ByVal endRow As Integer)

    Dim i As Integer
    For i = startRow To endRow
        ws.Range(colLetter & i).Formula = "=WORKDAY(A" & i & ",B" & i & ")"
        ws.Range(colLetter & i).NumberFormat = "yyyy/m/d"
    Next i
End Sub

Private Sub BatchInsertNetworkdaysFormula( _
    ByVal ws As Worksheet, _
    ByVal colLetter As String, _
    ByVal startRow As Integer, _
    ByVal endRow As Integer)

    Dim i As Integer
    For i = startRow To endRow
        ws.Range(colLetter & i).Formula = "=NETWORKDAYS(A" & i & ",D" & i & ")"
    Next i
End Sub

Private Sub BatchInsertNextMonthFirstWorkdayFormula( _
    ByVal ws As Worksheet, _
    ByVal colLetter As String, _
    ByVal startRow As Integer, _
    ByVal endRow As Integer)

    Dim i As Integer
    For i = startRow To endRow
        ws.Range(colLetter & i).Formula = _
            "=WORKDAY(DATE(YEAR(A" & i & "),MONTH(A" & i & ")+1,1)-1,1)"
        ws.Range(colLetter & i).NumberFormat = "yyyy/m/d"
    Next i
End Sub
