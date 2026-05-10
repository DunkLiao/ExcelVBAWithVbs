Option Explicit
'*************************************************************************************
'模組名稱: WeekendDateFormatting
'功能說明: 以條件格式標示週末日期範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyWeekendDateFormatting()
    Dim ws As Worksheet
    Dim calendarRange As Range
    Dim fc As FormatCondition

    On Error GoTo ErrHandler

    Set ws = GetOrCreateWeekendSheet("週末日期格式範例")
    ws.Cells.Clear
    Call FillWeekendDateData(ws)

    Set calendarRange = ws.Range("A2:G6")
    calendarRange.FormatConditions.Delete

    Set fc = calendarRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=WEEKDAY(A2,2)>5")
    With fc
        .Interior.Color = RGB(217, 225, 242)
        .Font.Color = RGB(47, 85, 151)
        .Font.Bold = True
    End With

    ws.Columns("A:G").ColumnWidth = 12
    MsgBox "週末日期條件格式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立週末日期條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillWeekendDateData(ByVal ws As Worksheet)
    Dim startDate As Date
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim dayOffset As Long

    ws.Range("A1:G1").Value = Array("星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日")
    startDate = Date - Weekday(Date, vbMonday) + 1
    dayOffset = 0

    For rowIndex = 2 To 6
        For columnIndex = 1 To 7
            ws.Cells(rowIndex, columnIndex).Value = startDate + dayOffset
            ws.Cells(rowIndex, columnIndex).NumberFormat = "m/d"
            dayOffset = dayOffset + 1
        Next columnIndex
    Next rowIndex
End Sub

Private Function GetOrCreateWeekendSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWeekendSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWeekendSheet Is Nothing Then
        Set GetOrCreateWeekendSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateWeekendSheet.Name = sheetName
    End If
End Function
