Attribute VB_Name = "BatchDateFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchDateFormulas
'功能說明: 批次填入日期公式，包含 YEAR、MONTH、DAY、DATEDIF、EDATE、EOMONTH
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchDateFormulas()
    Call CreateDateFormulaExample
End Sub

' 建立日期公式批次填入示範
Sub CreateDateFormulaExample()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreateDateSheet(ThisWorkbook, "日期公式示範")
    Call FillEmployeeHireDates(ws)
    Call BatchEnterDateExtractFormulas(ws)
    Call BatchEnterDateCalcFormulas(ws)

    ws.Columns("A:J").AutoFit
    ws.Activate
    MsgBox "日期公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次填入 YEAR / MONTH / DAY 拆解公式
Private Sub BatchEnterDateExtractFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("C1").Value = "年份(YEAR)"
    ws.Range("D1").Value = "月份(MONTH)"
    ws.Range("E1").Value = "日期(DAY)"
    ws.Range("C1:E1").Font.Bold = True

    For i = 2 To lastRow
        ws.Cells(i, 3).Formula = "=YEAR(B" & i & ")"
        ws.Cells(i, 4).Formula = "=MONTH(B" & i & ")"
        ws.Cells(i, 5).Formula = "=DAY(B" & i & ")"
    Next i
End Sub

' 批次填入 DATEDIF / EDATE / EOMONTH 計算公式
Private Sub BatchEnterDateCalcFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("F1").Value = "年資(年)"
    ws.Range("G1").Value = "年資(月)"
    ws.Range("H1").Value = "到職6個月日"
    ws.Range("I1").Value = "當月最後一天"
    ws.Range("J1").Value = "試用期結束(90天)"
    ws.Range("F1:J1").Font.Bold = True

    For i = 2 To lastRow
        ' DATEDIF 計算年資（年/月）
        ws.Cells(i, 6).Formula = "=DATEDIF(B" & i & ",TODAY(),""Y"")"
        ws.Cells(i, 7).Formula = "=DATEDIF(B" & i & ",TODAY(),""M"")"
        ' EDATE：到職後6個月
        ws.Cells(i, 8).Formula = "=EDATE(B" & i & ",6)"
        ws.Cells(i, 8).NumberFormat = "yyyy/mm/dd"
        ' EOMONTH：到職月份最後一天
        ws.Cells(i, 9).Formula = "=EOMONTH(B" & i & ",0)"
        ws.Cells(i, 9).NumberFormat = "yyyy/mm/dd"
        ' 到職日 + 90 天
        ws.Cells(i, 10).Formula = "=B" & i & "+90"
        ws.Cells(i, 10).NumberFormat = "yyyy/mm/dd"
    Next i
End Sub

' 填入員工到職日資料
Private Sub FillEmployeeHireDates(ByVal ws As Worksheet)
    Dim i As Integer
    Dim names(1 To 8) As String
    Dim hireDates(1 To 8) As Date

    names(1) = "王小明" : hireDates(1) = CDate("2020/03/15")
    names(2) = "李美玲" : hireDates(2) = CDate("2019/07/01")
    names(3) = "陳大偉" : hireDates(3) = CDate("2022/11/20")
    names(4) = "林志豪" : hireDates(4) = CDate("2018/01/10")
    names(5) = "張雅惠" : hireDates(5) = CDate("2023/04/05")
    names(6) = "黃建宏" : hireDates(6) = CDate("2021/09/30")
    names(7) = "蔡佳慧" : hireDates(7) = CDate("2024/02/14")
    names(8) = "吳宗翰" : hireDates(8) = CDate("2017/06/25")

    ws.Range("A1").Value = "員工姓名"
    ws.Range("B1").Value = "到職日期"
    ws.Range("A1:B1").Font.Bold = True

    For i = 1 To 8
        ws.Cells(i + 1, 1).Value = names(i)
        ws.Cells(i + 1, 2).Value = hireDates(i)
        ws.Cells(i + 1, 2).NumberFormat = "yyyy/mm/dd"
    Next i
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateDateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateDateSheet = ws
End Function