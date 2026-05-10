Option Explicit
'*************************************************************************************
'模組名稱: ExpiringDateFormatting
'功能說明: 建立合約到期日提醒的條件格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyExpiringDateFormatting()
    Dim ws As Worksheet
    Dim dateRange As Range
    Dim fc As FormatCondition

    On Error GoTo ErrHandler

    Set ws = GetOrCreateExpiringDateSheet("到期提醒格式範例")
    ws.Cells.Clear
    Call FillExpiringDateData(ws)

    Set dateRange = ws.Range("C2:C9")
    dateRange.FormatConditions.Delete

    Set fc = dateRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(C2<>"""",C2<TODAY())")
    With fc
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With

    Set fc = dateRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(C2>=TODAY(),C2<=TODAY()+30)")
    With fc
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
    End With

    ws.Columns("A:C").AutoFit
    MsgBox "逾期與 30 天內到期提醒已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立到期日條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillExpiringDateData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("供應商", "合約類型", "到期日")
    ws.Range("A2:C2").Value = Array("大同資訊", "維護", Date - 12)
    ws.Range("A3:C3").Value = Array("宏達顧問", "顧問", Date + 7)
    ws.Range("A4:C4").Value = Array("信義物流", "配送", Date + 19)
    ws.Range("A5:C5").Value = Array("安華保全", "保全", Date + 45)
    ws.Range("A6:C6").Value = Array("遠景科技", "授權", Date + 90)
    ws.Range("A7:C7").Value = Array("聯合清潔", "清潔", Date - 3)
    ws.Range("A8:C8").Value = Array("永信設備", "租賃", Date + 28)
    ws.Range("A9:C9").Value = Array("華通電信", "線路", Date + 62)
    ws.Range("C2:C9").NumberFormat = "yyyy/m/d"
End Sub

Private Function GetOrCreateExpiringDateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateExpiringDateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateExpiringDateSheet Is Nothing Then
        Set GetOrCreateExpiringDateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateExpiringDateSheet.Name = sheetName
    End If
End Function
