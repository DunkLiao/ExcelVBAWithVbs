Option Explicit

' 使用圖示集建立條件式格式，標示績效高低。
Public Sub ApplyIconSetFormattingExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim iconCondition As IconSetCondition

    Set ws = GetOrCreateConditionWorksheet("圖示集格式範例")
    ws.Cells.Clear
    Call FillIconSetData(ws)

    Set targetRange = ws.Range("C2:C8")
    targetRange.FormatConditions.Delete
    Set iconCondition = targetRange.FormatConditions.AddIconSetCondition

    With iconCondition
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
        .ShowIconOnly = False
        .IconCriteria(2).Type = xlConditionValueNumber
        .IconCriteria(2).Value = 70
        .IconCriteria(3).Type = xlConditionValueNumber
        .IconCriteria(3).Value = 90
    End With

    ws.Columns("A:C").AutoFit
    MsgBox "圖示集條件式格式已套用完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "套用條件式格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillIconSetData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("員工", "部門", "達成率")
    ws.Range("A2:C2").Value = Array("王小明", "業務", 95)
    ws.Range("A3:C3").Value = Array("陳美華", "業務", 82)
    ws.Range("A4:C4").Value = Array("林志強", "業務", 68)
    ws.Range("A5:C5").Value = Array("張雅婷", "客服", 91)
    ws.Range("A6:C6").Value = Array("黃建國", "客服", 74)
    ws.Range("A7:C7").Value = Array("李佳玲", "客服", 63)
    ws.Range("A8:C8").Value = Array("周信宏", "管理", 88)
    ws.Range("C2:C8").NumberFormat = "0"
End Sub

Private Function GetOrCreateConditionWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateConditionWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateConditionWorksheet Is Nothing Then
        Set GetOrCreateConditionWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateConditionWorksheet.Name = sheetName
    End If
End Function