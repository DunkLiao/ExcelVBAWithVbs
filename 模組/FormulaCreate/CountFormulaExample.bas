Option Explicit
Attribute VB_Name = "CountFormulaExample"
'*************************************************************************************
'模組名稱: CountFormulaExample
'功能說明: 批次產生計數公式，包含 COUNT、COUNTA、COUNTBLANK、COUNTIF、COUNTIFS 函數
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub CreateCountFormulaExample()
    Dim ws As Worksheet
    Dim i As Integer

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear

    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "業績"
    ws.Range("C1").Value = "狀態"

    Dim depts(1 To 8) As String
    Dim sales(1 To 8) As Variant
    Dim statuses(1 To 8) As String

    depts(1) = "業務一部" : sales(1) = 85000  : statuses(1) = "達標"
    depts(2) = "業務二部" : sales(2) = 42000  : statuses(2) = "未達標"
    depts(3) = "業務一部" : sales(3) = 93000  : statuses(3) = "達標"
    depts(4) = "業務三部" : sales(4) = ""     : statuses(4) = "待確認"
    depts(5) = "業務二部" : sales(5) = 61000  : statuses(5) = "達標"
    depts(6) = "業務三部" : sales(6) = 27000  : statuses(6) = "未達標"
    depts(7) = "業務一部" : sales(7) = 110000 : statuses(7) = "達標"
    depts(8) = "業務三部" : sales(8) = 55000  : statuses(8) = "達標"

    For i = 1 To 8
        ws.Cells(i + 1, 1).Value = depts(i)
        ws.Cells(i + 1, 2).Value = sales(i)
        ws.Cells(i + 1, 3).Value = statuses(i)
    Next i

    ws.Range("E1").Value = "公式說明"
    ws.Range("F1").Value = "結果"

    ws.Range("E2").Value = "COUNT（數值個數）"
    ws.Range("F2").Formula = "=COUNT(B2:B9)"

    ws.Range("E3").Value = "COUNTA（非空個數）"
    ws.Range("F3").Formula = "=COUNTA(B2:B9)"

    ws.Range("E4").Value = "COUNTBLANK（空白個數）"
    ws.Range("F4").Formula = "=COUNTBLANK(B2:B9)"

    ws.Range("E5").Value = "COUNTIF（業務一部筆數）"
    ws.Range("F5").Formula = "=COUNTIF(A2:A9,""業務一部"")"

    ws.Range("E6").Value = "COUNTIF（達標筆數）"
    ws.Range("F6").Formula = "=COUNTIF(C2:C9,""達標"")"

    ws.Range("E7").Value = "COUNTIFS（業務一部且達標）"
    ws.Range("F7").Formula = "=COUNTIFS(A2:A9,""業務一部"",C2:C9,""達標"")"

    ws.Range("E8").Value = "COUNTIFS（業績>50000且達標）"
    ws.Range("F8").Formula = "=COUNTIFS(B2:B9,"">50000"",C2:C9,""達標"")"

    ws.Columns("A:F").AutoFit

    MsgBox "計數公式範例已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "建立計數公式失敗"
End Sub