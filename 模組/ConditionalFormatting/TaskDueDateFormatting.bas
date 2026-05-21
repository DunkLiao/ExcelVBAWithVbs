Option Explicit
Attribute VB_Name = "TaskDueDateFormatting"
'*************************************************************************************
'模組名稱: TaskDueDateFormatting
'功能說明: 根據任務截止日期自動套用條件式格式，標示逾期、即將到期與正常任務
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestTaskDueDateFormatting()
    Call ApplyTaskDueDateFormatting("任務截止日期格式")
End Sub

Sub ApplyTaskDueDateFormatting(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateTDFSheet(sheetName)
    ws.Cells.Clear

    Call FillTaskData(ws)
    Call SetDueDateConditionalFormat(ws)

    ws.Columns.AutoFit
    MsgBox "任務截止日期條件式格式已套用完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "套用格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillTaskData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("任務名稱", "負責人", "截止日期")

    Dim today As Date
    today = Date

    ws.Range("A2:C2").Value = Array("撰寫報告", "張三", today - 3)
    ws.Range("A3:C3").Value = Array("更新資料庫", "李四", today + 2)
    ws.Range("A4:C4").Value = Array("客戶簡報", "王五", today + 8)
    ws.Range("A5:C5").Value = Array("系統測試", "趙六", today - 1)
    ws.Range("A6:C6").Value = Array("年度稽核", "陳七", today + 30)

    ws.Range("C2:C6").NumberFormat = "yyyy/m/d"
End Sub

Private Sub SetDueDateConditionalFormat(ByVal ws As Worksheet)
    Dim dataRange As Range
    Set dataRange = ws.Range("A2:C6")
    dataRange.FormatConditions.Delete

    ' 逾期：截止日期小於今日，紅色背景
    Dim fcOverdue As FormatCondition
    Set fcOverdue = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=$C2<TODAY()")

    With fcOverdue
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With

    ' 即將到期：截止日期在今日後 7 天內，黃色背景
    Dim fcSoon As FormatCondition
    Set fcSoon = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND($C2>=TODAY(),$C2<=TODAY()+7)")

    With fcSoon
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With

    ' 正常：截止日期在 7 天後，綠色背景
    Dim fcNormal As FormatCondition
    Set fcNormal = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=$C2>TODAY()+7")

    With fcNormal
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With
End Sub

Private Function GetOrCreateTDFSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateTDFSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateTDFSheet Is Nothing Then
        Set GetOrCreateTDFSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateTDFSheet.Name = sheetName
    End If
End Function
