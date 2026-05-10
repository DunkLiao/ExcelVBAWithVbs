Option Explicit
'*************************************************************************************
'模組名稱: PercentileBandFormatting
'功能說明: 建立第 90 百分位與第 10 百分位條件格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyPercentileBandFormatting()
    Dim ws As Worksheet
    Dim valueRange As Range
    Dim fc As FormatCondition

    On Error GoTo ErrHandler

    Set ws = GetOrCreatePercentileSheet("百分位格式範例")
    ws.Cells.Clear
    Call FillPercentileData(ws)

    Set valueRange = ws.Range("B2:B21")
    valueRange.FormatConditions.Delete

    Set fc = valueRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=PERCENTILE.INC($B$2:$B$21,0.9)")
    With fc
        .Interior.Color = RGB(112, 173, 71)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    Set fc = valueRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="=PERCENTILE.INC($B$2:$B$21,0.1)")
    With fc
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ws.Columns("A:B").AutoFit
    MsgBox "第 90 與第 10 百分位條件格式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立百分位條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillPercentileData(ByVal ws As Worksheet)
    Dim index As Long
    Dim values As Variant

    values = Array(72, 85, 61, 94, 77, 83, 58, 90, 69, 75, 88, 66, 97, 54, 80, 73, 91, 63, 86, 59)
    ws.Range("A1:B1").Value = Array("序號", "品質分數")

    For index = LBound(values) To UBound(values)
        ws.Cells(index + 2, 1).Value = index + 1
        ws.Cells(index + 2, 2).Value = values(index)
    Next index
End Sub

Private Function GetOrCreatePercentileSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreatePercentileSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreatePercentileSheet Is Nothing Then
        Set GetOrCreatePercentileSheet = ThisWorkbook.Worksheets.Add
        GetOrCreatePercentileSheet.Name = sheetName
    End If
End Function
