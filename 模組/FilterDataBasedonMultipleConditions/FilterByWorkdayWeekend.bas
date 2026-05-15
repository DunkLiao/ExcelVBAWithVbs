Attribute VB_Name = "FilterByWorkdayWeekend"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByWorkdayWeekend
'功能說明: 依據欄 A 日期將作用中工作表篩選為工作日或週末資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunFilterByWorkdayWeekend()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim userChoice As String
    Dim showWeekend As Boolean
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim visibleCount As Long

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    userChoice = Trim$(InputBox( _
        Prompt:="請選擇篩選條件:" & vbCrLf & _
                "1 = 工作日" & vbCrLf & _
                "2 = 週末", _
        Title:="工作日與週末篩選", _
        Default:="1"))

    If userChoice = "" Then Exit Sub

    Select Case userChoice
        Case "1"
            showWeekend = False
        Case "2"
            showWeekend = True
        Case Else
            MsgBox "請輸入 1 或 2。", vbExclamation, "提示"
            Exit Sub
    End Select

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "欄 A 沒有可篩選的日期資料。", vbInformation, "提示"
        Exit Sub
    End If

    ws.Rows("2:" & lastRow).Hidden = False

    For rowIndex = 2 To lastRow
        If IsDate(ws.Cells(rowIndex, 1).Value) Then
            If ShouldShowDate(CDate(ws.Cells(rowIndex, 1).Value), showWeekend) Then
                visibleCount = visibleCount + 1
            Else
                ws.Rows(rowIndex).Hidden = True
            End If
        Else
            ws.Rows(rowIndex).Hidden = True
        End If
    Next rowIndex

    MsgBox "已顯示 " & visibleCount & " 筆" & IIf(showWeekend, "週末", "工作日") & "資料。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "篩選工作日或週末資料時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function ShouldShowDate(ByVal targetDate As Date, ByVal showWeekend As Boolean) As Boolean
    Dim weekdayIndex As Long
    Dim isWeekend As Boolean

    weekdayIndex = Weekday(targetDate, vbMonday)
    isWeekend = (weekdayIndex > 5)
    ShouldShowDate = (isWeekend = showWeekend)
End Function
