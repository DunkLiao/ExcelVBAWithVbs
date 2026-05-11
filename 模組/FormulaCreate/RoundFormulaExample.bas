Attribute VB_Name = "RoundFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: RoundFormulaExample
'功能說明: 示範 ROUND、ROUNDUP、ROUNDDOWN、MROUND 四捨五入相關公式的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口
Sub TestRoundFormulas()
    Call CreateRoundFormulaSheet("四捨五入公式")
End Sub

' 建立四捨五入公式範例工作表
' sheetName: 工作表名稱
Sub CreateRoundFormulaSheet(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Call FillRoundData(ws)
    Call InsertRoundFormulas(ws)
    ws.Columns("A:E").AutoFit
    MsgBox "四捨五入公式範例已建立完成！", vbInformation, "完成"
End Sub

' 填入來源數值資料與標題
Private Sub FillRoundData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "原始數值"
    ws.Range("B1").Value = "ROUND(2位)"
    ws.Range("C1").Value = "ROUNDUP(2位)"
    ws.Range("D1").Value = "ROUNDDOWN(2位)"
    ws.Range("E1").Value = "MROUND(0.5)"

    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Range("A2").Value = 3.14159
    ws.Range("A3").Value = 2.71828
    ws.Range("A4").Value = -1.5678
    ws.Range("A5").Value = 9.995
    ws.Range("A6").Value = 12.3456
    ws.Range("A7").Value = 0.12345
    ws.Range("A8").Value = 100.449
    ws.Range("A9").Value = 100.451
    ws.Range("A10").Value = -7.777
End Sub

' 批次插入四捨五入公式
Private Sub InsertRoundFormulas(ByVal ws As Worksheet)
    Dim i As Integer

    For i = 2 To 10
        ws.Cells(i, 2).Formula = "=ROUND(A" & i & ",2)"
        ws.Cells(i, 3).Formula = "=ROUNDUP(A" & i & ",2)"
        ws.Cells(i, 4).Formula = "=ROUNDDOWN(A" & i & ",2)"
        ws.Cells(i, 5).Formula = "=MROUND(A" & i & ",0.5)"
    Next i

    ws.Range("A2:E10").NumberFormat = "0.00"
End Sub