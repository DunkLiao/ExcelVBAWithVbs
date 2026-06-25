Option Explicit
Attribute VB_Name = "PivotWithMacroButton"
'*************************************************************************************
'模組名稱: PivotWithMacroButton
'功能說明: 在工作表中放置按鈕，點擊後自動更新或重新整理樞紐分析表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestPivotWithButton()
    Call CreatePivotWithMacroButton
End Sub

' 建立樞紐分析表並加入更新按鈕
Sub CreatePivotWithMacroButton()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim btn As Button
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheet("樞紐按鈕範例")
    ws.Cells.Clear

    ' 建立範例資料
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "銷售量"
    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 100
    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 150
    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 200
    ws.Range("A5").Value = "產品A"
    ws.Range("B5").Value = 130
    ws.Range("A6").Value = "產品B"
    ws.Range("B6").Value = 180

    Set dataRange = ws.Range("A1:B6")

    ' 建立樞紐分析表
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=ws.Range("D1"), _
        TableName:="銷售樞紐")

    With pt
        .PivotFields("產品").Orientation = xlRowField
        .PivotFields("銷售量").Orientation = xlDataField
    End With

    ' 建立按鈕以觸發樞紐重新整理
    Set btn = ws.Buttons.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top - 30, _
        Width:=120, _
        Height:=25)
    With btn
        .Caption = "重新整理樞紐"
        .OnAction = "RefreshPivotButton"
        .Font.Size = 10
    End With

    MsgBox "已建立樞紐分析表及更新按鈕。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 按鈕觸發的重新整理程序
Sub RefreshPivotButton()
    Dim ws As Worksheet
    Dim pt As PivotTable

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If ws.PivotTables.Count = 0 Then
        MsgBox "此工作表沒有樞紐分析表。", vbExclamation, "提示"
        Exit Sub
    End If

    On Error Resume Next
    Set pt = ws.PivotTables(1)
    pt.PivotCache.Refresh
    If Err.Number = 0 Then
        MsgBox "樞紐分析表已重新整理。", vbInformation, "完成"
    Else
        MsgBox "重新整理失敗：" & Err.Description, vbCritical, "錯誤"
    End If
    On Error GoTo 0
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
