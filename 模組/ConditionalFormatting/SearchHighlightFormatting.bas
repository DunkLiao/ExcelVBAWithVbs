Option Explicit
Attribute VB_Name = "SearchHighlightFormatting"
'*************************************************************************************
'模組名稱: SearchHighlightFormatting
'功能說明: 依使用者輸入的關鍵字，在指定範圍內搜尋並以格式化醒目提示符合的儲存格
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestSearchHighlight()
    Call SearchAndHighlightFormatting
End Sub

' 搜尋並以格式化醒目提示符合的儲存格
Sub SearchAndHighlightFormatting()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim searchRange As Range
    Dim searchText As String
    Dim fc As FormatCondition

    Set ws = GetOrCreateWorksheet("搜尋醒目提示")
    ws.Cells.Clear

    ' 建立範例資料
    ws.Range("A1").Value = "編號"
    ws.Range("B1").Value = "產品名稱"
    ws.Range("C1").Value = "庫存量"
    ws.Range("A2").Value = "P001"
    ws.Range("B2").Value = "筆記型電腦"
    ws.Range("C2").Value = 50
    ws.Range("A3").Value = "P002"
    ws.Range("B3").Value = "桌上型電腦"
    ws.Range("C3").Value = 30
    ws.Range("A4").Value = "P003"
    ws.Range("B4").Value = "平板電腦"
    ws.Range("C4").Value = 80
    ws.Range("A5").Value = "P004"
    ws.Range("B5").Value = "智慧型手機"
    ws.Range("C5").Value = 120
    ws.Range("A6").Value = "P005"
    ws.Range("B6").Value = "平板電腦Pro"
    ws.Range("C6").Value = 45

    ' 要求使用者輸入搜尋關鍵字
    searchText = InputBox("請輸入要搜尋的關鍵字：", "關鍵字搜尋", "電腦")
    If searchText = "" Then Exit Sub

    ' 清除原有的條件格式
    Set searchRange = ws.UsedRange
    searchRange.FormatConditions.Delete

    ' 加入醒目提示條件格式
    Set fc = searchRange.FormatConditions.Add( _
        Type:=xlTextString, _
        String:=searchText, _
        TextOperator:=xlContains)

    With fc
        .Interior.Color = RGB(255, 255, 0)
        .Font.Bold = True
        .Font.Color = RGB(192, 0, 0)
    End With

    ' 自動調整欄寬
    ws.Columns.AutoFit

    MsgBox "已將包含「" & searchText & "」的儲存格以黃底紅字醒目提示。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
