Attribute VB_Name = "BatchNestedIFFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchNestedIFFormulas
'功能說明: 批次在指定欄位寫入巢狀 IF 公式，依分數自動評定等第
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口
Sub TestBatchNestedIFFormulas()
    Call CreateNestedIFGradeSheet("成績等第評定")
End Sub

' 建立成績工作表並批次寫入巢狀 IF 等第公式
' sheetName: 工作表名稱
Sub CreateNestedIFGradeSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Call FillScoreData(ws)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Call BatchWriteGradeFormula(ws, 2, lastRow, 3, 2)

    ws.Columns("A:C").AutoFit
    MsgBox "批次巢狀 IF 公式寫入完成！共處理 " & (lastRow - 1) & " 筆資料。", _
           vbInformation, "完成"
End Sub

' 批次寫入巢狀 IF 等第公式
' ws         : 目標工作表
' startRow   : 起始資料列
' endRow     : 結束資料列
' formulaCol : 公式輸出欄號
' scoreCol   : 分數來源欄號
Sub BatchWriteGradeFormula( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal formulaCol As Integer, _
    ByVal scoreCol As Integer)

    Dim r As Long
    Dim scoreRef As String
    Dim formulaText As String
    Dim scoreColLetter As String

    scoreColLetter = Split(ws.Cells(1, scoreCol).Address(True, False), "$")(0)

    For r = startRow To endRow
        scoreRef = scoreColLetter & CStr(r)
        formulaText = "=IF(" & scoreRef & ">=90,""A""," & _
                        "IF(" & scoreRef & ">=80,""B""," & _
                        "IF(" & scoreRef & ">=70,""C""," & _
                        "IF(" & scoreRef & ">=60,""D"",""F""))))"
        ws.Cells(r, formulaCol).Formula = formulaText
    Next r
End Sub

' 填入成績範例資料
Private Sub FillScoreData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "學號"
    ws.Range("B1").Value = "分數"
    ws.Range("C1").Value = "等第"
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(70, 130, 180)
    ws.Range("A1:C1").Font.Color = RGB(255, 255, 255)

    Dim scores(1 To 12) As Integer
    scores(1) = 95: scores(2) = 82: scores(3) = 76: scores(4) = 65
    scores(5) = 58: scores(6) = 91: scores(7) = 73: scores(8) = 88
    scores(9) = 44: scores(10) = 100: scores(11) = 69: scores(12) = 50

    Dim i As Integer
    For i = 1 To 12
        ws.Cells(i + 1, 1).Value = "S" & Format(i, "00")
        ws.Cells(i + 1, 2).Value = scores(i)
    Next i
End Sub