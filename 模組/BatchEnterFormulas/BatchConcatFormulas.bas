Option Explicit
Attribute VB_Name = "BatchConcatFormulas"
'*************************************************************************************
'模組名稱: BatchConcatFormulas
'功能說明: 批次產生 CONCAT、CONCATENATE 及 & 運算子字串串接公式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub BatchConcatFormulasExample()
    Dim ws As Worksheet
    Dim i As Integer

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear

    ws.Range("A1").Value = "姓"
    ws.Range("B1").Value = "名"
    ws.Range("C1").Value = "部門"
    ws.Range("D1").Value = "分機"

    Dim surnames(1 To 6) As String
    Dim givenNames(1 To 6) As String
    Dim depts(1 To 6) As String
    Dim exts(1 To 6) As String

    surnames(1) = "王"  : givenNames(1) = "大明" : depts(1) = "業務部" : exts(1) = "1001"
    surnames(2) = "李"  : givenNames(2) = "小華" : depts(2) = "財務部" : exts(2) = "2003"
    surnames(3) = "張"  : givenNames(3) = "志豪" : depts(3) = "研發部" : exts(3) = "3012"
    surnames(4) = "陳"  : givenNames(4) = "雅婷" : depts(4) = "人資部" : exts(4) = "4025"
    surnames(5) = "林"  : givenNames(5) = "俊傑" : depts(5) = "業務部" : exts(5) = "1008"
    surnames(6) = "黃"  : givenNames(6) = "美玲" : depts(6) = "行銷部" : exts(6) = "5016"

    For i = 1 To 6
        ws.Cells(i + 1, 1).Value = surnames(i)
        ws.Cells(i + 1, 2).Value = givenNames(i)
        ws.Cells(i + 1, 3).Value = depts(i)
        ws.Cells(i + 1, 4).Value = exts(i)
    Next i

    ws.Range("F1").Value = "公式方式"
    ws.Range("G1").Value = "串接結果"

    ws.Range("F2").Value = "CONCATENATE（姓+名）"
    ws.Range("G2").Formula = "=CONCATENATE(A2,B2)"

    ws.Range("F3").Value = "& 運算子（姓+名）"
    ws.Range("G3").Formula = "=A3&B3"

    ws.Range("F4").Value = "CONCAT（姓+名+部門）"
    ws.Range("G4").Formula = "=CONCAT(A4,B4,""-"",C4)"

    ws.Range("F5").Value = "& 加分機格式"
    ws.Range("G5").Formula = "=A5&B5&""（""&C5&"" 分機 ""&D5&"")"""

    ws.Range("F6").Value = "CONCAT 多欄合併"
    ws.Range("G6").Formula = "=CONCAT(A6:D6)"

    ws.Range("F7").Value = "TEXTJOIN（以逗號分隔）"
    ws.Range("G7").Formula = "=TEXTJOIN("","",TRUE,A7,B7,C7,D7)"

    ws.Columns("A:G").AutoFit

    MsgBox "字串串接公式範例已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "建立串接公式失敗"
End Sub