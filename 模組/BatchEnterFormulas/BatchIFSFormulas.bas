Attribute VB_Name = "BatchIFSFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchIFSFormulas
'功能說明: 批次在工作表儲存格中輸入IFS多條件判斷公式，並自動填充欄位
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestBatchIFSFormulas()
    Call CreateIFSFormulaExample
End Sub

' 建立IFS公式批次輸入範例
Sub CreateIFSFormulaExample()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreateIFSSheet(ThisWorkbook, "IFS公式範例")
    Call FillStudentBaseData(ws)
    Call BatchEnterIFSGrade(ws)
    Call BatchEnterIFSBonus(ws)

    ws.Columns("A:H").AutoFit
    ws.Activate
    MsgBox "IFS多條件判斷公式已批次輸入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次輸入IFS成績等第公式
Private Sub BatchEnterIFSGrade(ByVal ws As Worksheet)
    Dim i As Integer
    ws.Range("F1").Value = "成績等第(IFS)"
    ws.Range("F1").Font.Bold = True
    For i = 2 To 11
        ws.Cells(i, 6).Formula = _
            "=IFS(" & ws.Cells(i, 5).Address(False, False) & ">=90,""A""," & _
            ws.Cells(i, 5).Address(False, False) & ">=80,""B""," & _
            ws.Cells(i, 5).Address(False, False) & ">=70,""C""," & _
            ws.Cells(i, 5).Address(False, False) & ">=60,""D""," & _
            "TRUE,""F"")"
    Next i
End Sub

' 批次輸入IFS獎金公式
Private Sub BatchEnterIFSBonus(ByVal ws As Worksheet)
    Dim i As Integer
    ws.Range("G1").Value = "獎金級別(IFS)"
    ws.Range("H1").Value = "獎金金額"
    ws.Range("G1:H1").Font.Bold = True
    For i = 2 To 11
        ws.Cells(i, 7).Formula = _
            "=IFS(" & ws.Cells(i, 5).Address(False, False) & ">=90,""特優""," & _
            ws.Cells(i, 5).Address(False, False) & ">=80,""優良""," & _
            ws.Cells(i, 5).Address(False, False) & ">=70,""普通""," & _
            ws.Cells(i, 5).Address(False, False) & ">=60,""尚可""," & _
            "TRUE,""不合格"")"
        ws.Cells(i, 8).Formula = _
            "=IFS(" & ws.Cells(i, 5).Address(False, False) & ">=90,3000," & _
            ws.Cells(i, 5).Address(False, False) & ">=80,2000," & _
            ws.Cells(i, 5).Address(False, False) & ">=70,1000," & _
            ws.Cells(i, 5).Address(False, False) & ">=60,500," & _
            "TRUE,0)"
        ws.Cells(i, 8).NumberFormat = "#,##0"
    Next i
End Sub

' 填入學生成績基礎資料
Private Sub FillStudentBaseData(ByVal ws As Worksheet)
    Dim i As Integer
    Dim names(1 To 10) As String
    Dim scores(1 To 10) As Integer
    Dim subjects(1 To 10) As String

    ws.Range("A1").Value = "學號"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "科目"
    ws.Range("D1").Value = "考試日期"
    ws.Range("E1").Value = "分數"
    ws.Range("A1:E1").Font.Bold = True

    names(1) = "王小明" : names(2) = "陳美玲" : names(3) = "林大偉" : names(4) = "黃雅琪"
    names(5) = "李建宏" : names(6) = "張惠芳" : names(7) = "吳俊傑" : names(8) = "劉怡君"
    names(9) = "蔡宗翰" : names(10) = "許雅婷"

    scores(1) = 95 : scores(2) = 78 : scores(3) = 52 : scores(4) = 86
    scores(5) = 43 : scores(6) = 98 : scores(7) = 67 : scores(8) = 31
    scores(9) = 74 : scores(10) = 89

    subjects(1) = "數學" : subjects(2) = "英語" : subjects(3) = "數學"
    subjects(4) = "英語" : subjects(5) = "數學" : subjects(6) = "英語"
    subjects(7) = "數學" : subjects(8) = "英語" : subjects(9) = "數學"
    subjects(10) = "英語"

    For i = 1 To 10
        ws.Cells(i + 1, 1).Value = "S" & Format(i, "000")
        ws.Cells(i + 1, 2).Value = names(i)
        ws.Cells(i + 1, 3).Value = subjects(i)
        ws.Cells(i + 1, 4).Value = DateSerial(2025, 5, i)
        ws.Cells(i + 1, 4).NumberFormat = "yyyy/mm/dd"
        ws.Cells(i + 1, 5).Value = scores(i)
    Next i
End Sub

' 取得或建立工作表並清除內容
Private Function GetOrCreateIFSSheet(ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateIFSSheet = ws
End Function
