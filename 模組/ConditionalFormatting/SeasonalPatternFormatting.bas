Attribute VB_Name = "SeasonalPatternFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: SeasonalPatternFormatting
'功能說明: 以VBA依據季節性模式（春夏秋冬）自動套用不同背景色條件式格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestSeasonalPatternFormatting()
    Call ApplySeasonalFormatting
End Sub

' 套用季節性模式條件式格式
Sub ApplySeasonalFormatting()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim fc As FormatCondition
    Dim lngLastRow As Long

    On Error GoTo ErrHandler
    Set ws = GetOrCreateSeasonSheet(ThisWorkbook, "季節性格式範例")
    Call FillSeasonData(ws)

    lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set dataRange = ws.Range("A2:C" & lngLastRow)
    dataRange.FormatConditions.Delete

    ' 春季（3~5月）：淺綠色背景
    Set fc = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND($B2>=3,$B2<=5)")
    fc.Interior.Color = RGB(198, 239, 206)

    ' 夏季（6~8月）：淺黃色背景
    Set fc = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND($B2>=6,$B2<=8)")
    fc.Interior.Color = RGB(255, 242, 204)

    ' 秋季（9~11月）：淺橙色背景
    Set fc = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND($B2>=9,$B2<=11)")
    fc.Interior.Color = RGB(255, 217, 179)

    ' 冬季（12月、1月、2月）：淺藍色背景
    Set fc = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=OR($B2=12,$B2<=2)")
    fc.Interior.Color = RGB(189, 215, 238)

    ws.Range("E1").Value = "季節色彩說明"
    ws.Range("E1").Font.Bold = True
    ws.Range("E2").Value = "春季（3~5月）"
    ws.Range("E2").Interior.Color = RGB(198, 239, 206)
    ws.Range("E3").Value = "夏季（6~8月）"
    ws.Range("E3").Interior.Color = RGB(255, 242, 204)
    ws.Range("E4").Value = "秋季（9~11月）"
    ws.Range("E4").Interior.Color = RGB(255, 217, 179)
    ws.Range("E5").Value = "冬季（12、1、2月）"
    ws.Range("E5").Interior.Color = RGB(189, 215, 238)

    ws.Columns("A:E").AutoFit
    ws.Activate
    MsgBox "季節性模式條件式格式已套用完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 填入月份資料
Private Sub FillSeasonData(ByVal ws As Worksheet)
    Dim i As Integer
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "月份"
    ws.Range("C1").Value = "銷售額"
    ws.Range("A1:C1").Font.Bold = True

    For i = 1 To 12
        ws.Cells(i + 1, 1).Value = DateSerial(2025, i, 1)
        ws.Cells(i + 1, 1).NumberFormat = "yyyy/mm/dd"
        ws.Cells(i + 1, 2).Value = i
        ws.Cells(i + 1, 3).Value = 80000 + (i * 3000)
    Next i
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表並清除內容
Private Function GetOrCreateSeasonSheet(ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSeasonSheet = ws
End Function
