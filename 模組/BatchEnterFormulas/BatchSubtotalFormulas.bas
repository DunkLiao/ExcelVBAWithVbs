Attribute VB_Name = "BatchSubtotalFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchSubtotalFormulas
'功能說明: 批次輸入 SUBTOTAL 公式（含 SUM、AVERAGE、COUNT、MAX、MIN 等多種彙總類型）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestBatchSubtotalFormulas()
    Call CreateBatchSubtotalExample
End Sub

Sub CreateBatchSubtotalExample()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim wsName As String
    wsName = "SUBTOTAL範例"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName
    
    ' 標題列
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "銷售額"
    ws.Range("C1").Value = "成本"
    ws.Range("D1").Value = "毛利"
    ws.Range("E1").Value = "客戶數"
    ws.Range("A1:E1").Font.Bold = True
    
    ' 範例資料
    Dim sales As Variant
    sales = Array(150000, 220000, 180000, 250000, 300000, 210000, 170000, 280000)
    Dim costs As Variant
    costs = Array(90000, 130000, 110000, 140000, 180000, 120000, 100000, 160000)
    
    For i = 1 To 8
        ws.Cells(i + 1, 1).Value = i & "月"
        ws.Cells(i + 1, 2).Value = sales(i - 1)
        ws.Cells(i + 1, 3).Value = costs(i - 1)
        ws.Cells(i + 1, 4).Value = sales(i - 1) - costs(i - 1)
        ws.Cells(i + 1, 5).Value = Int(Rnd * 50) + 10
    Next i
    
    ws.Range("B2:D9").NumberFormat = "#,##0"
    lastRow = 9
    
    ' 批次輸入 SUBTOTAL 公式
    ' 加總
    Dim sumRow As Long
    sumRow = lastRow + 2
    ws.Cells(sumRow, 1).Value = "SUBTOTAL 彙總"
    ws.Cells(sumRow, 1).Font.Bold = True
    
    ws.Cells(sumRow + 1, 1).Value = "加總 (9)"
    ws.Cells(sumRow + 1, 2).Formula = "=SUBTOTAL(9,B2:B" & lastRow & ")"
    ws.Cells(sumRow + 1, 3).Formula = "=SUBTOTAL(9,C2:C" & lastRow & ")"
    ws.Cells(sumRow + 1, 4).Formula = "=SUBTOTAL(9,D2:D" & lastRow & ")"
    
    ' 平均值
    ws.Cells(sumRow + 2, 1).Value = "平均值 (1)"
    ws.Cells(sumRow + 2, 2).Formula = "=SUBTOTAL(1,B2:B" & lastRow & ")"
    ws.Cells(sumRow + 2, 3).Formula = "=SUBTOTAL(1,C2:C" & lastRow & ")"
    ws.Cells(sumRow + 2, 4).Formula = "=SUBTOTAL(1,D2:D" & lastRow & ")"
    
    ' 最大值
    ws.Cells(sumRow + 3, 1).Value = "最大值 (4)"
    ws.Cells(sumRow + 3, 2).Formula = "=SUBTOTAL(4,B2:B" & lastRow & ")"
    ws.Cells(sumRow + 3, 3).Formula = "=SUBTOTAL(4,C2:C" & lastRow & ")"
    ws.Cells(sumRow + 3, 4).Formula = "=SUBTOTAL(4,D2:D" & lastRow & ")"
    
    ' 最小值
    ws.Cells(sumRow + 4, 1).Value = "最小值 (5)"
    ws.Cells(sumRow + 4, 2).Formula = "=SUBTOTAL(5,B2:B" & lastRow & ")"
    ws.Cells(sumRow + 4, 3).Formula = "=SUBTOTAL(5,C2:C" & lastRow & ")"
    ws.Cells(sumRow + 4, 4).Formula = "=SUBTOTAL(5,D2:D" & lastRow & ")"
    
    ' 計數
    ws.Cells(sumRow + 5, 1).Value = "計數 (2)"
    ws.Cells(sumRow + 5, 2).Formula = "=SUBTOTAL(2,B2:B" & lastRow & ")"
    ws.Cells(sumRow + 5, 5).Formula = "=SUBTOTAL(2,E2:E" & lastRow & ")"
    
    ' 標準差
    ws.Cells(sumRow + 6, 1).Value = "標準差 (7)"
    ws.Cells(sumRow + 6, 2).Formula = "=SUBTOTAL(7,B2:B" & lastRow & ")"
    
    ' 格式化彙總區域
    ws.Range("B" & (sumRow + 1) & ":D" & (sumRow + 4)).NumberFormat = "#,##0"
    ws.Range("B" & (sumRow + 5)).NumberFormat = "0"
    ws.Range("B" & (sumRow + 6)).NumberFormat = "#,##0"
    
    ws.Columns.AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "SUBTOTAL 公式批次輸入完成！" & vbCrLf & _
           "函數編號對照：1=AVERAGE, 2=COUNT, 4=MAX, 5=MIN, 7=STDEV, 9=SUM", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "批次輸入公式時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
