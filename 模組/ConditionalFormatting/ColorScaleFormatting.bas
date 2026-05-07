Attribute VB_Name = "ColorScaleFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ColorScaleFormatting
'功能說明: 使用VBA對指定範圍套用色階條件式格式設定
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口
Sub TestColorScaleFormatting()
    Call ApplyColorScaleFormatting
End Sub

' 對資料範圍套用三色色階條件式格式
Sub ApplyColorScaleFormatting()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim cf As ColorScale
    Dim barRange As Range

    Set ws = GetOrCreateSheet(ThisWorkbook, "色階格式範例")
    Call FillGradeData(ws)

    Set dataRange = ws.Range("B2:D6")
    dataRange.FormatConditions.Delete

    ' 新增三色色階格式（低=紅, 中=黃, 高=綠）
    Set cf = dataRange.FormatConditions.AddColorScale(ColorScaleType:=3)

    With cf.ColorScaleCriteria(1)
        .Type = xlConditionValueLowestValue
        .FormatColor.Color = RGB(255, 0, 0)
    End With

    With cf.ColorScaleCriteria(2)
        .Type = xlConditionValuePercentile
        .Value = 50
        .FormatColor.Color = RGB(255, 255, 0)
    End With

    With cf.ColorScaleCriteria(3)
        .Type = xlConditionValueHighestValue
        .FormatColor.Color = RGB(0, 255, 0)
    End With

    ' 示範資料橫條格式
    Set barRange = ws.Range("E2:E6")
    ws.Range("E1").Value = "業績達成率(%)"
    ws.Range("E2").Value = 45
    ws.Range("E3").Value = 78
    ws.Range("E4").Value = 92
    ws.Range("E5").Value = 60
    ws.Range("E6").Value = 85
    barRange.FormatConditions.Delete
    barRange.FormatConditions.AddDatabar

    ws.Columns("A:E").AutoFit
    ws.Activate
    MsgBox "色階條件式格式已套用完成！", vbInformation, "完成"
End Sub

' 填入成績範例資料
Private Sub FillGradeData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "學生"
    ws.Range("B1").Value = "國文"
    ws.Range("C1").Value = "英文"
    ws.Range("D1").Value = "數學"
    ws.Range("A2").Value = "陳小明"
    ws.Range("B2").Value = 78
    ws.Range("C2").Value = 85
    ws.Range("D2").Value = 92
    ws.Range("A3").Value = "李小華"
    ws.Range("B3").Value = 55
    ws.Range("C3").Value = 60
    ws.Range("D3").Value = 48
    ws.Range("A4").Value = "王小美"
    ws.Range("B4").Value = 90
    ws.Range("C4").Value = 88
    ws.Range("D4").Value = 95
    ws.Range("A5").Value = "張小強"
    ws.Range("B5").Value = 65
    ws.Range("C5").Value = 72
    ws.Range("D5").Value = 68
    ws.Range("A6").Value = "林小芬"
    ws.Range("B6").Value = 82
    ws.Range("C6").Value = 77
    ws.Range("D6").Value = 80
    ws.Columns("A:D").AutoFit
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
