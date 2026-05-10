Option Explicit
'*************************************************************************************
'模組名稱: BlankErrorCellFormatting
'功能說明: 建立空白儲存格與錯誤值的條件格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyBlankErrorCellFormatting()
    Dim ws As Worksheet
    Dim checkRange As Range
    Dim fc As FormatCondition

    On Error GoTo ErrHandler

    Set ws = GetOrCreateBlankErrorSheet("空白錯誤格式範例")
    ws.Cells.Clear
    Call FillBlankErrorData(ws)

    Set checkRange = ws.Range("B2:D8")
    checkRange.FormatConditions.Delete

    Set fc = checkRange.FormatConditions.Add(Type:=xlBlanksCondition)
    With fc
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = RGB(89, 89, 89)
    End With

    Set fc = checkRange.FormatConditions.Add(Type:=xlErrorsCondition)
    With fc
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With

    ws.Columns("A:D").AutoFit
    MsgBox "空白與錯誤值條件格式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立空白錯誤條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillBlankErrorData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("專案", "一月", "二月", "三月")
    ws.Range("A2:D2").Value = Array("專案A", 12, 15, 18)
    ws.Range("A3:D3").Value = Array("專案B", 9, "", 14)
    ws.Range("A4:D4").Value = Array("專案C", 20, 22, "")
    ws.Range("A5:D5").Value = Array("專案D", 8, 11, 13)
    ws.Range("A6:D6").Value = Array("專案E", "", 17, 19)
    ws.Range("A7").Value = "專案F"
    ws.Range("B7").Formula = "=10/0"
    ws.Range("C7").Value = 16
    ws.Range("D7").Value = 21
    ws.Range("A8").Value = "專案G"
    ws.Range("B8").Value = 14
    ws.Range("C8").Formula = "=NA()"
    ws.Range("D8").Value = 18
End Sub

Private Function GetOrCreateBlankErrorSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateBlankErrorSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateBlankErrorSheet Is Nothing Then
        Set GetOrCreateBlankErrorSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateBlankErrorSheet.Name = sheetName
    End If
End Function
