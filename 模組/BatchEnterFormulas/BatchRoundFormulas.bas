Attribute VB_Name = "BatchRoundFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchRoundFormulas
'功能說明: 批次在 Excel 儲存格中建立 ROUND、ROUNDUP、ROUNDDOWN
'          以及 MROUND、INT、TRUNC 等捨入相關公式的範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub TestBatchRoundFormulas()
    Call CreateBatchRoundFormulas("批次捨入公式範例")
End Sub

Sub CreateBatchRoundFormulas(ByVal sheetName As String)
    Dim ws      As Worksheet
    Dim i       As Long
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    Call FillRoundSourceData(ws)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        ws.Cells(i, 3).Formula = "=ROUND(A" & i & ",2)"
        ws.Cells(i, 4).Formula = "=ROUNDUP(A" & i & ",2)"
        ws.Cells(i, 5).Formula = "=ROUNDDOWN(A" & i & ",2)"
        ws.Cells(i, 6).Formula = "=MROUND(A" & i & ",0.5)"
        ws.Cells(i, 7).Formula = "=INT(A" & i & ")"
        ws.Cells(i, 8).Formula = "=TRUNC(A" & i & ",1)"
    Next i

    ws.Range("C1").Value = "ROUND(,2)"
    ws.Range("D1").Value = "ROUNDUP(,2)"
    ws.Range("E1").Value = "ROUNDDOWN(,2)"
    ws.Range("F1").Value = "MROUND(,0.5)"
    ws.Range("G1").Value = "INT"
    ws.Range("H1").Value = "TRUNC(,1)"
    ws.Range("C1:H1").Interior.Color = RGB(189, 215, 238)
    ws.Range("C2:H" & lastRow).NumberFormat = "0.00"
    ws.Columns("A:H").AutoFit
    MsgBox "批次捨入公式已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillRoundSourceData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "原始數值"
    ws.Range("B1").Value = "說明"
    ws.Range("A1:B1").Font.Bold = True

    Dim values(9) As Double
    Dim labels(9) As String

    values(0) = 3.14159  : labels(0) = "圓周率 pi"
    values(1) = 2.71828  : labels(1) = "自然數 e"
    values(2) = 123.456  : labels(2) = "一般金額"
    values(3) = 99.999   : labels(3) = "接近整數"
    values(4) = -7.654   : labels(4) = "負數"
    values(5) = 0.005    : labels(5) = "極小正數"
    values(6) = 1000.001 : labels(6) = "大數微差"
    values(7) = 55.555   : labels(7) = "對稱小數"
    values(8) = -0.125   : labels(8) = "負小數"
    values(9) = 8.76543  : labels(9) = "多位小數"

    Dim i As Integer
    For i = 0 To 9
        ws.Cells(i + 2, 1).Value = values(i)
        ws.Cells(i + 2, 2).Value = labels(i)
    Next i
End Sub
