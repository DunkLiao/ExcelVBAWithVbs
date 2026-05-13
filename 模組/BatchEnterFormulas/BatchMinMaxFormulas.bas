Attribute VB_Name = "BatchMinMaxFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchMinMaxFormulas
'功能說明: 批次在指定欄位末尾輸入 MIN / MAX / MEDIAN 統計公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub BatchEnterMinMaxFormulas()
    On Error GoTo ErrHandler
    Dim ws         As Worksheet
    Dim lastRow    As Long
    Dim lastCol    As Long
    Dim summaryRow As Long
    Dim c          As Long
    Dim colAddr    As String
    Dim summaryRng As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then
        MsgBox "工作表資料不足，請確認至少有標題列與一列數值資料。", vbExclamation, "提示"
        Exit Sub
    End If
    summaryRow = lastRow + 2
    ws.Cells(summaryRow - 1, 1).Value = "統計列"
    ws.Cells(summaryRow - 1, 1).Font.Bold = True
    ws.Cells(summaryRow,     1).Value = "最大值"
    ws.Cells(summaryRow + 1, 1).Value = "最小值"
    ws.Cells(summaryRow + 2, 1).Value = "中位數"
    For c = 2 To lastCol
        colAddr = ws.Cells(2, c).Address(False, True) & ":" & _
                  ws.Cells(lastRow, c).Address(False, True)
        If IsNumeric(ws.Cells(2, c).Value) Then
            ws.Cells(summaryRow,     c).Formula = "=MAX(" & colAddr & ")"
            ws.Cells(summaryRow + 1, c).Formula = "=MIN(" & colAddr & ")"
            ws.Cells(summaryRow + 2, c).Formula = "=MEDIAN(" & colAddr & ")"
            ws.Cells(summaryRow,     c).NumberFormat = ws.Cells(2, c).NumberFormat
            ws.Cells(summaryRow + 1, c).NumberFormat = ws.Cells(2, c).NumberFormat
            ws.Cells(summaryRow + 2, c).NumberFormat = ws.Cells(2, c).NumberFormat
        End If
    Next c
    Set summaryRng = ws.Range(ws.Cells(summaryRow, 1), ws.Cells(summaryRow + 2, lastCol))
    With summaryRng
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
    End With
    ws.Columns.AutoFit
    MsgBox "已批次建立 MAX / MIN / MEDIAN 公式，位於第 " & summaryRow & " 至 " & summaryRow + 2 & " 列。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Public Sub CreateMinMaxSampleData()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("最大最小值公式範例")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "最大最小值公式範例"
    End If
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("姓名", "國文", "數學", "英文")
    ws.Range("A2:D2").Value = Array("張小明", 85, 92, 78)
    ws.Range("A3:D3").Value = Array("李美華", 76, 68, 91)
    ws.Range("A4:D4").Value = Array("王大同", 93, 88, 83)
    ws.Range("A5:D5").Value = Array("陳志偉", 62, 77, 95)
    ws.Range("A6:D6").Value = Array("林雅芳", 89, 95, 72)
    ws.Range("A1:D1").Font.Bold = True
    ws.Columns.AutoFit
    ws.Activate
    MsgBox "測試資料已建立，請執行 BatchEnterMinMaxFormulas 建立統計公式。", vbInformation, "完成"
End Sub

