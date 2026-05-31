Attribute VB_Name = "AggregateFormulaExample"
Option Explicit

'*************************************************************************************
'模組名稱: AggregateFormulaExample
'功能說明: 使用 AGGREGATE 函數忽略錯誤值進行統計
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub InsertAggregateFormulas()
    '在指定欄位插入 AGGREGATE 函數，忽略錯誤值與隱藏列進行統計
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim resultRow As Long

    Set ws = ThisWorkbook.Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    resultRow = lastRow + 2

    ' AGGREGATE 函數號：1=AVERAGE, 9=SUM, 2=COUNT，選項6=忽略錯誤值
    ws.Cells(resultRow, 1).Value = "AGGREGATE 統計結果"
    ws.Cells(resultRow, 2).Formula = _
        "=AGGREGATE(9,6,B2:B" & lastRow & ")"
    ws.Cells(resultRow + 1, 2).Formula = _
        "=AGGREGATE(1,6,B2:B" & lastRow & ")"
    ws.Cells(resultRow + 2, 2).Formula = _
        "=AGGREGATE(2,6,B2:B" & lastRow & ")"

    ws.Cells(resultRow, 1).Font.Bold = True
    MsgBox "AGGREGATE 公式已插入第 " & resultRow & " 列！", vbInformation
End Sub
