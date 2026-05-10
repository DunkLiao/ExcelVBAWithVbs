Option Explicit
'*************************************************************************************
'模組名稱: UniqueValueFormatting
'功能說明: 建立唯一值提示的條件格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyUniqueValueFormatting()
    Dim ws As Worksheet
    Dim codeRange As Range
    Dim fc As UniqueValues

    On Error GoTo ErrHandler

    Set ws = GetOrCreateUniqueValueSheet("唯一值格式範例")
    ws.Cells.Clear
    Call FillUniqueValueData(ws)

    Set codeRange = ws.Range("B2:B14")
    codeRange.FormatConditions.Delete

    Set fc = codeRange.FormatConditions.AddUniqueValues
    fc.DupeUnique = xlUnique
    With fc
        .Interior.Color = RGB(189, 215, 238)
        .Font.Color = RGB(31, 78, 121)
        .Font.Bold = True
    End With

    ws.Columns("A:C").AutoFit
    MsgBox "唯一值條件格式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立唯一值條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillUniqueValueData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("序號", "客戶代碼", "客戶名稱")
    ws.Range("A2:C2").Value = Array(1, "C001", "北區商行")
    ws.Range("A3:C3").Value = Array(2, "C002", "東門企業")
    ws.Range("A4:C4").Value = Array(3, "C003", "南港材料")
    ws.Range("A5:C5").Value = Array(4, "C002", "東門企業")
    ws.Range("A6:C6").Value = Array(5, "C004", "中山食品")
    ws.Range("A7:C7").Value = Array(6, "C005", "大安科技")
    ws.Range("A8:C8").Value = Array(7, "C006", "信義通路")
    ws.Range("A9:C9").Value = Array(8, "C003", "南港材料")
    ws.Range("A10:C10").Value = Array(9, "C007", "松山製造")
    ws.Range("A11:C11").Value = Array(10, "C008", "內湖服務")
    ws.Range("A12:C12").Value = Array(11, "C004", "中山食品")
    ws.Range("A13:C13").Value = Array(12, "C009", "士林零售")
    ws.Range("A14:C14").Value = Array(13, "C010", "萬華工程")
End Sub

Private Function GetOrCreateUniqueValueSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateUniqueValueSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateUniqueValueSheet Is Nothing Then
        Set GetOrCreateUniqueValueSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateUniqueValueSheet.Name = sheetName
    End If
End Function
