Attribute VB_Name = "BatchIFFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchIFFormulas
'功能說明: 批次在工作表中填入 IF、巢狀 IF 判斷公式，依成績自動標示等第
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchIFFormulas()
    Call CreateIFFormulaExample
End Sub

' 建立 IF 公式批次填入示範
Sub CreateIFFormulaExample()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreateIFSheet(ThisWorkbook, "IF公式示範")
    Call FillScoreBaseData(ws)
    Call BatchEnterGradeFormulas(ws)
    Call BatchEnterPassFailFormulas(ws)

    ws.Columns("A:G").AutoFit
    ws.Activate
    MsgBox "IF 判斷公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次填入巢狀 IF 等第公式 (A/B/C/D/F)
Private Sub BatchEnterGradeFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    ' 平均分數在 F 欄，等第放 G 欄
    ws.Range("F1").Value = "平均"
    ws.Range("G1").Value = "等第"
    ws.Range("F1:G1").Font.Bold = True

    For i = 2 To 11
        ' 先計算平均
        ws.Cells(i, 6).Formula = "=AVERAGE(" & _
            ws.Cells(i, 2).Address(False, False) & ":" & _
            ws.Cells(i, 5).Address(False, False) & ")"
        ws.Cells(i, 6).NumberFormat = "0.0"

        ' 巢狀 IF：90+ A / 80+ B / 70+ C / 60+ D / 不及格 F
        ws.Cells(i, 7).Formula = "=IF(" & ws.Cells(i, 6).Address(False, False) & _
            ">=90,""A"",IF(" & ws.Cells(i, 6).Address(False, False) & _
            ">=80,""B"",IF(" & ws.Cells(i, 6).Address(False, False) & _
            ">=70,""C"",IF(" & ws.Cells(i, 6).Address(False, False) & _
            ">=60,""D"",""F""))))"
    Next i
End Sub

' 批次填入及格/不及格判斷公式
Private Sub BatchEnterPassFailFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    Dim j As Integer
    ' 在每科後面（H~K欄）加上通過/未通過
    ws.Range("H1").Value = "國文狀態"
    ws.Range("I1").Value = "英文狀態"
    ws.Range("J1").Value = "數學狀態"
    ws.Range("K1").Value = "自然狀態"
    ws.Range("H1:K1").Font.Bold = True

    For i = 2 To 11
        For j = 0 To 3
            ws.Cells(i, 8 + j).Formula = "=IF(" & _
                ws.Cells(i, 2 + j).Address(False, False) & _
                ">=60,""通過"",""未通過"")"
        Next j
    Next i
    ws.Columns("A:K").AutoFit
End Sub

' 填入基礎成績資料
Private Sub FillScoreBaseData(ByVal ws As Worksheet)
    Dim i As Integer
    Dim names(1 To 10) As String
    Dim scores(1 To 10, 1 To 4) As Integer

    names(1) = "王小明" : names(2) = "李美玲" : names(3) = "陳大偉" : names(4) = "林志豪"
    names(5) = "張雅惠" : names(6) = "黃建宏" : names(7) = "蔡佳慧" : names(8) = "吳宗翰"
    names(9) = "鄭麗娟" : names(10) = "許志遠"

    scores(1, 1) = 92 : scores(1, 2) = 88 : scores(1, 3) = 95 : scores(1, 4) = 91
    scores(2, 1) = 75 : scores(2, 2) = 70 : scores(2, 3) = 68 : scores(2, 4) = 73
    scores(3, 1) = 55 : scores(3, 2) = 60 : scores(3, 3) = 45 : scores(3, 4) = 58
    scores(4, 1) = 83 : scores(4, 2) = 87 : scores(4, 3) = 80 : scores(4, 4) = 85
    scores(5, 1) = 40 : scores(5, 2) = 35 : scores(5, 3) = 50 : scores(5, 4) = 42
    scores(6, 1) = 98 : scores(6, 2) = 95 : scores(6, 3) = 92 : scores(6, 4) = 97
    scores(7, 1) = 62 : scores(7, 2) = 68 : scores(7, 3) = 65 : scores(7, 4) = 60
    scores(8, 1) = 30 : scores(8, 2) = 28 : scores(8, 3) = 35 : scores(8, 4) = 32
    scores(9, 1) = 78 : scores(9, 2) = 82 : scores(9, 3) = 75 : scores(9, 4) = 80
    scores(10, 1) = 68 : scores(10, 2) = 72 : scores(10, 3) = 70 : scores(10, 4) = 65

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
Private Function GetOrCreateIFSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateIFSheet = ws
End Function