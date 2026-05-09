Attribute VB_Name = "BatchRankFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchRankFormulas
'功能說明: 批次填入 RANK 排名公式，對多欄成績自動產生名次與百分比排名
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchRankFormulas()
    Call CreateRankFormulaExample
End Sub

' 建立排名公式批次填入示範
Sub CreateRankFormulaExample()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreateRankSheet(ThisWorkbook, "排名公式示範")
    Call FillSalesPerformanceData(ws)
    Call BatchEnterRankFormulas(ws)
    Call BatchEnterPercentRankFormulas(ws)

    ws.Columns("A:H").AutoFit
    ws.Activate
    MsgBox "排名公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次填入 RANK 名次公式（由大到小）
Private Sub BatchEnterRankFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("D1").Value = "業績名次(RANK)"
    ws.Range("E1").Value = "客戶數名次(RANK)"
    ws.Range("D1:E1").Font.Bold = True

    ' 固定範圍參考（加上 $ 絕對參照）
    Dim salesRange As String
    Dim custRange As String
    salesRange = "$B$2:$B$" & lastRow
    custRange = "$C$2:$C$" & lastRow

    For i = 2 To lastRow
        ' RANK 第3參數 0 = 由大到小
        ws.Cells(i, 4).Formula = "=RANK(B" & i & "," & salesRange & ",0)"
        ws.Cells(i, 5).Formula = "=RANK(C" & i & "," & custRange & ",0)"
    Next i
End Sub

' 批次填入百分比排名與超過平均標記
Private Sub BatchEnterPercentRankFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("F1").Value = "業績%排名"
    ws.Range("G1").Value = "超過平均?"
    ws.Range("H1").Value = "排名標籤"
    ws.Range("F1:H1").Font.Bold = True

    Dim salesRange As String
    salesRange = "$B$2:$B$" & lastRow

    For i = 2 To lastRow
        ' PERCENTRANK.INC 百分比排名
        ws.Cells(i, 6).Formula = "=PERCENTRANK.INC(" & salesRange & ",B" & i & ")"
        ws.Cells(i, 6).NumberFormat = "0.0%"
        ' IF：超過平均則標示是
        ws.Cells(i, 7).Formula = "=IF(B" & i & ">AVERAGE(" & salesRange & "),""是"",""否"")"
        ' 排名標籤文字
        ws.Cells(i, 8).Formula = "=""第""&D" & i & "&""名"""
    Next i
End Sub

' 填入業務績效資料
Private Sub FillSalesPerformanceData(ByVal ws As Worksheet)
    Dim i As Integer
    Dim names(1 To 8) As String
    Dim sales(1 To 8) As Long
    Dim customers(1 To 8) As Integer

    names(1) = "業務一" : sales(1) = 1250000 : customers(1) = 35
    names(2) = "業務二" : sales(2) = 980000  : customers(2) = 28
    names(3) = "業務三" : sales(3) = 1580000 : customers(3) = 42
    names(4) = "業務四" : sales(4) = 760000  : customers(4) = 22
    names(5) = "業務五" : sales(5) = 2100000 : customers(5) = 55
    names(6) = "業務六" : sales(6) = 1380000 : customers(6) = 38
    names(7) = "業務七" : sales(7) = 890000  : customers(7) = 25
    names(8) = "業務八" : sales(8) = 1720000 : customers(8) = 47

    ws.Range("A1").Value = "業務姓名"
    ws.Range("B1").Value = "業績金額"
    ws.Range("C1").Value = "客戶數"
    ws.Range("A1:C1").Font.Bold = True

    For i = 1 To 8
        ws.Cells(i + 1, 1).Value = names(i)
        ws.Cells(i + 1, 2).Value = sales(i)
        ws.Cells(i + 1, 2).NumberFormat = "#,##0"
        ws.Cells(i + 1, 3).Value = customers(i)
    Next i
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateRankSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateRankSheet = ws
End Function