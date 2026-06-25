Option Explicit
Attribute VB_Name = "BatchTrimFormulas"
'*************************************************************************************
'模組名稱: BatchTrimFormulas
'功能說明: 批次在指定欄位輸入 TRIM、CLEAN、UPPER、LOWER、PROPER 等文字處理公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestBatchTrimFormulas()
    Call BatchEnterTrimFormulas
End Sub

' 批次輸入文字處理公式
Sub BatchEnterTrimFormulas()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = GetOrCreateWorksheet("文字處理公式")
    ws.Cells.Clear

    ' 建立範例資料（含多餘空白與大小寫不一致）
    ws.Range("A1").Value = "原始資料"
    ws.Range("B1").Value = "TRIM結果"
    ws.Range("C1").Value = "CLEAN結果"
    ws.Range("D1").Value = "UPPER結果"
    ws.Range("E1").Value = "LOWER結果"
    ws.Range("F1").Value = "PROPER結果"

    ws.Range("A2").Value = "  apple  "
    ws.Range("A3").Value = "  banana" & vbLf & "  "
    ws.Range("A4").Value = "   HELLO world"
    ws.Range("A5").Value = "  mICROSOFT excel  "
    ws.Range("A6").Value = "  john smith  "

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 批次輸入公式
    For i = 2 To lastRow
        ws.Cells(i, 2).Formula = "=TRIM(A" & i & ")"
        ws.Cells(i, 3).Formula = "=CLEAN(A" & i & ")"
        ws.Cells(i, 4).Formula = "=UPPER(A" & i & ")"
        ws.Cells(i, 5).Formula = "=LOWER(A" & i & ")"
        ws.Cells(i, 6).Formula = "=PROPER(A" & i & ")"
    Next i

    ws.Columns.AutoFit

    MsgBox "已在 " & (lastRow - 1) & " 列中批次輸入文字處理公式。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
