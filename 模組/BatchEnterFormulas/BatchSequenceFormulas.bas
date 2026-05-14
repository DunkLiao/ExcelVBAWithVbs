Attribute VB_Name = "BatchSequenceFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchSequenceFormulas
'功能說明: 批次在 Excel 儲存格中輸入 SEQUENCE、ROW、COLUMN 等序列相關公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestBatchSequenceFormulas()
    Call CreateBatchSequenceExample
End Sub

' 建立批次序列公式範例
Sub CreateBatchSequenceExample()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateSeqWs(ThisWorkbook, "序列公式範例")
    ws.Cells.Clear

    ws.Range("A1").Value = "SEQUENCE 函數範例"
    ws.Range("A1").Font.Bold = True

    ws.Range("A2").Value = "1 到 10 的序列"
    ws.Range("B2").Formula = "=SEQUENCE(10)"

    ws.Range("A3").Value = "3 列 4 欄的矩陣序列"
    ws.Range("B3").Formula = "=SEQUENCE(3,4)"

    ws.Range("A4").Value = "從 5 開始步進 3"
    ws.Range("B4").Formula = "=SEQUENCE(8,1,5,3)"

    ws.Range("A6").Value = "ROW 序列（相容版）"
    ws.Range("A6").Font.Bold = True

    Dim i As Long
    For i = 1 To 10
        ws.Cells(6, i + 1).Formula = "=ROW()-5"
    Next i

    ws.Range("A7").Value = "COLUMN 序列（相容版）"
    For i = 1 To 10
        ws.Cells(7, i + 1).Formula = "=COLUMN()-1"
    Next i

    ws.Range("A9").Value = "連續日期序列（今日起 10 天）"
    ws.Range("A9").Font.Bold = True
    For i = 0 To 9
        ws.Cells(9, i + 2).Formula = "=TODAY()+" & i
        ws.Cells(9, i + 2).NumberFormat = "m/d"
    Next i

    ws.Range("A10").Value = "每月第一天（12 個月）"
    For i = 0 To 11
        ws.Cells(10, i + 2).Formula = "=DATE(YEAR(TODAY()),MONTH(TODAY())+" & i & ",1)"
        ws.Cells(10, i + 2).NumberFormat = "yyyy/mm"
    Next i

    ws.Range("A12").Value = "等差數列（步進公式）"
    ws.Range("A12").Font.Bold = True
    ws.Range("A13").Value = "起始值"
    ws.Range("B13").Value = 100
    ws.Range("A14").Value = "步進值"
    ws.Range("B14").Value = 15
    ws.Range("A15").Value = "等差序列"
    For i = 0 To 9
        ws.Cells(15, i + 2).Formula = "=$B$13+" & i & "*$B$14"
    Next i

    ws.Columns("A:L").AutoFit
    MsgBox "批次序列公式已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立序列公式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateSeqWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateSeqWs = ws
End Function