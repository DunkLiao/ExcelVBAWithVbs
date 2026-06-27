Attribute VB_Name = "BatchRunningBalanceFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchRunningBalanceFormulas
'功能說明: 批次輸入累計餘額公式（Running Balance），示範收支記帳的累積計算
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestBatchRunningBalanceFormulas()
    Call BatchEnterRunningBalanceFormulas
End Sub

Sub BatchEnterRunningBalanceFormulas()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("累計餘額公式")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "累計餘額公式"
    End If
    
    ws.Cells.Clear
    Call FillRunningBalanceData(ws)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ws.Range("D2").Formula = "=B2"
    
    For i = 3 To lastRow
        ws.Cells(i, 4).Formula = "=D" & i - 1 & "+B" & i
    Next i
    
    Dim q As String
    q = Chr(34)
    ws.Range("F2").Formula = "=IF(ISBLANK(D2)," & q & q & ",IF(D2<0," & q & "超支" & q & "," & q & "正常" & q & "))"
    
    For i = 3 To lastRow
        ws.Cells(i, 6).Formula = "=IF(ISBLANK(D" & i & ")," & q & q & ",IF(D" & i & "<0," & q & "超支" & q & "," & q & "正常" & q & "))"
    Next i
    
    ws.Range("A1:F1").Font.Bold = True
    
    With ws.Range("F2:F" & lastRow)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=" & q & "超支" & q
        .FormatConditions(1).Interior.Color = RGB(255, 200, 200)
        .FormatConditions(1).Font.Color = RGB(200, 0, 0)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=" & q & "正常" & q
        .FormatConditions(2).Interior.Color = RGB(200, 255, 200)
    End With
    
    ws.Columns("A:F").AutoFit
    ws.Activate
    
    MsgBox "累計餘額公式批次輸入完成！共 " & lastRow - 1 & " 筆。", vbInformation, "完成"
End Sub

Private Sub FillRunningBalanceData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "收支金額"
    ws.Range("C1").Value = "摘要"
    ws.Range("D1").Value = "累計餘額"
    ws.Range("E1").Value = "起始餘額"
    ws.Range("F1").Value = "狀態"
    
    ws.Range("A2").Value = "2024/1/1"
    ws.Range("B2").Value = 5000
    ws.Range("C2").Value = "起始存款"
    ws.Range("E2").Value = 0
    
    ws.Range("A3").Value = "2024/1/5"
    ws.Range("B3").Value = -1200
    ws.Range("C3").Value = "房租支出"
    
    ws.Range("A4").Value = "2024/1/10"
    ws.Range("B4").Value = 3000
    ws.Range("C4").Value = "薪資收入"
    
    ws.Range("A5").Value = "2024/1/15"
    ws.Range("B5").Value = -500
    ws.Range("C5").Value = "水電費"
    
    ws.Range("A6").Value = "2024/1/20"
    ws.Range("B6").Value = -800
    ws.Range("C6").Value = "餐飲支出"
    
    ws.Range("A7").Value = "2024/1/25"
    ws.Range("B7").Value = 4500
    ws.Range("C7").Value = "獎金收入"
    
    ws.Range("A8").Value = "2024/1/30"
    ws.Range("B8").Value = -2000
    ws.Range("C8").Value = "卡費支出"
    
    ws.Columns("A:C").AutoFit
End Sub
