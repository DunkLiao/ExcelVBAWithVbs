Option Explicit
'*************************************************************************************
'模組名稱: AboveBelowAverageFormatting
'功能說明: 建立高於平均與低於平均的條件格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyAboveBelowAverageFormatting()
    Dim ws As Worksheet
    Dim amountRange As Range
    Dim fc As AboveAverage

    On Error GoTo ErrHandler

    Set ws = GetOrCreateAverageSheet("平均值格式範例")
    ws.Cells.Clear
    Call FillAverageData(ws)

    Set amountRange = ws.Range("B2:B13")
    amountRange.FormatConditions.Delete

    Set fc = amountRange.FormatConditions.AddAboveAverage
    With fc
        .AboveBelow = xlAboveAverage
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(31, 78, 121)
        .Font.Bold = True
    End With

    Set fc = amountRange.FormatConditions.AddAboveAverage
    With fc
        .AboveBelow = xlBelowAverage
        .Interior.Color = RGB(252, 228, 214)
        .Font.Color = RGB(132, 60, 12)
    End With

    ws.Range("D1").Value = "平均銷售額"
    ws.Range("D2").Formula = "=AVERAGE(B2:B13)"
    ws.Columns("A:D").AutoFit
    MsgBox "高於平均與低於平均條件格式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立平均值條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillAverageData(ByVal ws As Worksheet)
    Dim monthNames As Variant
    Dim amounts As Variant
    Dim index As Long

    monthNames = Array("一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月")
    amounts = Array(125000, 98000, 142000, 117500, 156000, 88000, 133000, 101000, 149000, 96000, 171000, 109000)
    ws.Range("A1:B1").Value = Array("月份", "銷售額")

    For index = LBound(monthNames) To UBound(monthNames)
        ws.Cells(index + 2, 1).Value = monthNames(index)
        ws.Cells(index + 2, 2).Value = amounts(index)
    Next index
End Sub

Private Function GetOrCreateAverageSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateAverageSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateAverageSheet Is Nothing Then
        Set GetOrCreateAverageSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateAverageSheet.Name = sheetName
    End If
End Function
