Attribute VB_Name = "BatchCumulativeFormulas"
Option Explicit

'*************************************************************************************
'模組名稱: BatchCumulativeFormulas
'功能說明: 批次在指定欄位插入累積 SUM、AVERAGE、COUNT 公式
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub BatchInsertCumulativeFormulas()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "B 欄資料不足，請確認有資料後再執行！", vbExclamation
        Exit Sub
    End If

    ws.Range("C1").Value = "累積合計"
    ws.Range("D1").Value = "累積平均"
    ws.Range("E1").Value = "累積計數"
    ws.Range("C1:E1").Font.Bold = True

    For i = 2 To lastRow
        ws.Cells(i, "C").Formula = "=SUM($B$2:B" & i & ")"
        ws.Cells(i, "D").Formula = "=AVERAGE($B$2:B" & i & ")"
        ws.Cells(i, "E").Formula = "=COUNT($B$2:B" & i & ")"
    Next i

    ws.Columns("C:E").AutoFit
    MsgBox "已批次插入 " & (lastRow - 1) & " 列累積公式！", vbInformation
End Sub
