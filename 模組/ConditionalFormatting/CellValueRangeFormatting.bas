Attribute VB_Name = "CellValueRangeFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: CellValueRangeFormatting
'功能說明: 依據儲存格數值範圍（低中高）套用條件式格式，以不同顏色區分層級
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestCellValueRangeFormatting()
    Call ApplyCellValueRangeFormatting
End Sub

' 依數值範圍套用三層條件式格式（低/中/高）
Sub ApplyCellValueRangeFormatting()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim fc As FormatCondition

    Set ws = GetOrCreateSheet(ThisWorkbook, "數值範圍格式範例")
    Call FillScoreData(ws)

    Set targetRange = ws.Range("B2:E11")

    targetRange.FormatConditions.Delete

    ' 低分：0-59（紅色底）
    Set fc = targetRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlBetween, _
        Formula1:="0", _
        Formula2:="59")
    fc.Interior.Color = RGB(255, 199, 206)
    fc.Font.Color = RGB(156, 0, 6)
    fc.Font.Bold = True

    ' 中分：60-79（黃色底）
    Set fc = targetRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlBetween, _
        Formula1:="60", _
        Formula2:="79")
    fc.Interior.Color = RGB(255, 235, 156)
    fc.Font.Color = RGB(156, 101, 0)

    ' 高分：80-100（綠色底）
    Set fc = targetRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlBetween, _
        Formula1:="80", _
        Formula2:="100")
    fc.Interior.Color = RGB(198, 239, 206)
    fc.Font.Color = RGB(0, 97, 0)

    ws.Activate
    MsgBox "數值範圍條件式格式套用完成！" & vbCrLf & _
           "紅色 = 0-59 分，黃色 = 60-79 分，綠色 = 80-100 分", _
           vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "套用條件式格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 填入成績範例資料
Private Sub FillScoreData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "學生"
    ws.Range("B1").Value = "國文"
    ws.Range("C1").Value = "數學"
    ws.Range("D1").Value = "英文"
    ws.Range("E1").Value = "自然"
    ws.Range("A1:E1").Font.Bold = True

    Dim scores As Variant
    scores = Array( _
        Array("王小明", 92, 55, 78, 88), _
        Array("李大華", 67, 82, 45, 71), _
        Array("陳美玲", 85, 91, 87, 93), _
        Array("張志偉", 48, 62, 73, 55), _
        Array("林雅婷", 76, 68, 95, 80), _
        Array("吳建國", 59, 44, 62, 50), _
        Array("黃淑芬", 88, 77, 83, 90), _
        Array("楊明哲", 72, 85, 58, 67), _
        Array("蔡雅文", 95, 98, 92, 97), _
        Array("許俊傑", 61, 53, 70, 64) _
    )

    Dim i As Integer
    For i = 0 To 9
        ws.Cells(i + 2, 1).Value = scores(i)(0)
        ws.Cells(i + 2, 2).Value = scores(i)(1)
        ws.Cells(i + 2, 3).Value = scores(i)(2)
        ws.Cells(i + 2, 4).Value = scores(i)(3)
        ws.Cells(i + 2, 5).Value = scores(i)(4)
    Next i

    ws.Columns("A:E").AutoFit
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
