Attribute VB_Name = "ClearFormattingInRange"
Option Explicit
'*************************************************************************************
'模組名稱: ClearFormattingInRange
'功能說明: 清除指定範圍或整個工作表的儲存格格式設定
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口（建立含格式的範例資料）
Sub TestClearFormatting()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "清除格式範例")
    Call FillFormattedData(ws)
    MsgBox "已建立含格式的範例資料。" & vbCrLf & _
           "可執行以下程序進行測試：" & vbCrLf & _
           "  ClearAllFormatsInSheet     - 清除整張工作表格式" & vbCrLf & _
           "  ClearConditionalFormatsInSheet - 清除條件式格式" & vbCrLf & _
           "  ClearFormatInRange         - 清除指定範圍格式" & vbCrLf & _
           "  ResetRangeToDefault        - 重設選取範圍為預設格式", _
           vbInformation, "提示"
End Sub

' 清除整張工作表所有儲存格格式（保留資料內容）
Sub ClearAllFormatsInSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.UsedRange.ClearFormats
    MsgBox "已清除工作表「" & ws.Name & "」的所有格式！", vbInformation, "完成"
End Sub

' 清除整張工作表的條件式格式
Sub ClearConditionalFormatsInSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Cells.FormatConditions.Delete
    MsgBox "已清除工作表「" & ws.Name & "」的所有條件式格式！", vbInformation, "完成"
End Sub

' 清除指定範圍的儲存格格式（保留資料）
Sub ClearFormatInRange()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim rangeAddress As String

    Set ws = ActiveSheet
    rangeAddress = InputBox("請輸入要清除格式的儲存格範圍（例如：A1:D10）：", "輸入範圍")

    If rangeAddress = "" Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If

    On Error Resume Next
    Set targetRange = ws.Range(rangeAddress)
    On Error GoTo 0

    If targetRange Is Nothing Then
        MsgBox "範圍格式不正確，請重新輸入！", vbExclamation, "錯誤"
        Exit Sub
    End If

    targetRange.ClearFormats
    targetRange.FormatConditions.Delete
    MsgBox "已清除範圍 " & rangeAddress & " 的格式！", vbInformation, "完成"
End Sub

' 重設選取範圍為預設格式
Sub ResetRangeToDefault()
    Dim targetRange As Range
    Set targetRange = Selection

    If targetRange Is Nothing Then
        MsgBox "請先選取要重設的儲存格範圍！", vbExclamation, "警告"
        Exit Sub
    End If

    With targetRange
        .ClearFormats
        .Font.Name = "新細明體"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Italic = False
        .Font.Color = RGB(0, 0, 0)
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlNone
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .NumberFormat = "General"
    End With

    MsgBox "選取範圍已重設為預設格式！", vbInformation, "完成"
End Sub

' 填入含格式的範例資料
Private Sub FillFormattedData(ByVal ws As Worksheet)
    With ws.Range("A1:D1")
        .Value = Array("編號", "姓名", "部門", "薪資")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Range("A2").Value = 1
    ws.Range("B2").Value = "張三"
    ws.Range("C2").Value = "研發部"
    ws.Range("D2").Value = 60000

    ws.Range("A3").Value = 2
    ws.Range("B3").Value = "李四"
    ws.Range("C3").Value = "業務部"
    ws.Range("D3").Value = 55000
    ws.Range("A3:D3").Interior.Color = RGB(230, 240, 250)

    ws.Range("A4").Value = 3
    ws.Range("B4").Value = "王五"
    ws.Range("C4").Value = "行政部"
    ws.Range("D4").Value = 45000

    ws.Range("D2:D4").NumberFormat = "#,##0"
    ws.Range("A1:D4").Borders.LineStyle = xlContinuous
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
