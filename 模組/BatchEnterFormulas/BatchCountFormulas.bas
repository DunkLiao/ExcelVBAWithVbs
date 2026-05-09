Attribute VB_Name = "BatchCountFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchCountFormulas
'功能說明: 批次在多欄工作表中自動填入 COUNT、COUNTA、COUNTIF、COUNTIFS 統計公式
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchCountFormulas()
    Call CreateCountFormulaExample
End Sub

' 建立計數公式批次填入示範
Sub CreateCountFormulaExample()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreateSheet(ThisWorkbook, "計數公式示範")
    Call FillStudentScoreData(ws)
    Call BatchEnterCountFormulas(ws)

    ws.Columns("A:H").AutoFit
    ws.Activate
    MsgBox "計數公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次填入 COUNT / COUNTA / COUNTIF / COUNTIFS 公式
Private Sub BatchEnterCountFormulas(ByVal ws As Worksheet)
    Dim i As Integer

    ' 科目欄位：B~E (4科)
    For i = 2 To 5
        ' COUNT：計算數值筆數
        ws.Cells(12, i).Formula = "=COUNT(" & _
            ws.Cells(2, i).Address(False, False) & ":" & _
            ws.Cells(11, i).Address(False, False) & ")"
        ' COUNTA：含文字計算非空筆數
        ws.Cells(13, i).Formula = "=COUNTA(" & _
            ws.Cells(2, i).Address(False, False) & ":" & _
            ws.Cells(11, i).Address(False, False) & ")"
        ' COUNTIF：計算 >= 60 (及格人數)
        ws.Cells(14, i).Formula = "=COUNTIF(" & _
            ws.Cells(2, i).Address(False, False) & ":" & _
            ws.Cells(11, i).Address(False, False) & ","">=60"")"
        ' COUNTIF：計算 < 60 (不及格人數)
        ws.Cells(15, i).Formula = "=COUNTIF(" & _
            ws.Cells(2, i).Address(False, False) & ":" & _
            ws.Cells(11, i).Address(False, False) & ",""<60"")"
    Next i

    ' 標籤
    ws.Range("A12").Value = "數值筆數(COUNT)"
    ws.Range("A13").Value = "非空筆數(COUNTA)"
    ws.Range("A14").Value = "及格人數(>=60)"
    ws.Range("A15").Value = "不及格人數(<60)"
    ws.Range("A12:A15").Font.Bold = True
    ws.Range("B12:E15").Font.Bold = True
End Sub

' 填入學生成績示範資料
Private Sub FillStudentScoreData(ByVal ws As Worksheet)
    Dim i As Integer
    Dim names(1 To 10) As String
    Dim scores(1 To 10, 1 To 4) As Variant

    names(1) = "王小明" : names(2) = "李美玲" : names(3) = "陳大偉" : names(4) = "林志豪"
    names(5) = "張雅惠" : names(6) = "黃建宏" : names(7) = "蔡佳慧" : names(8) = "吳宗翰"
    names(9) = "鄭麗娟" : names(10) = "許志遠"

    scores(1, 1) = 85 : scores(1, 2) = 72 : scores(1, 3) = 91 : scores(1, 4) = 68
    scores(2, 1) = 90 : scores(2, 2) = 88 : scores(2, 3) = 75 : scores(2, 4) = 82
    scores(3, 1) = 55 : scores(3, 2) = 61 : scores(3, 3) = 48 : scores(3, 4) = 70
    scores(4, 1) = 78 : scores(4, 2) = 83 : scores(4, 3) = 89 : scores(4, 4) = 77
    scores(5, 1) = 45 : scores(5, 2) = 52 : scores(5, 3) = 60 : scores(5, 4) = 55
    scores(6, 1) = 92 : scores(6, 2) = 95 : scores(6, 3) = 88 : scores(6, 4) = 90
    scores(7, 1) = 67 : scores(7, 2) = 74 : scores(7, 3) = 70 : scores(7, 4) = 63
    scores(8, 1) = 38 : scores(8, 2) = 42 : scores(8, 3) = 50 : scores(8, 4) = 45
    scores(9, 1) = 80 : scores(9, 2) = 78 : scores(9, 3) = 82 : scores(9, 4) = 85
    scores(10, 1) = 73 : scores(10, 2) = 69 : scores(10, 3) = 76 : scores(10, 4) = 71

    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "國文"
    ws.Range("C1").Value = "英文"
    ws.Range("D1").Value = "數學"
    ws.Range("E1").Value = "自然"
    ws.Range("A1:E1").Font.Bold = True

    For i = 1 To 10
        ws.Cells(i + 1, 1).Value = names(i)
        ws.Cells(i + 1, 2).Value = scores(i, 1)
        ws.Cells(i + 1, 3).Value = scores(i, 2)
        ws.Cells(i + 1, 4).Value = scores(i, 3)
        ws.Cells(i + 1, 5).Value = scores(i, 4)
    Next i
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function