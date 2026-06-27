Attribute VB_Name = "PivotWithDynamicFieldSelection"
Option Explicit
'*************************************************************************************
'模組名稱: PivotWithDynamicFieldSelection
'功能說明: 建立樞紐分析表並允許使用者透過InputBox動態選取列欄位與值欄位
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestPivotWithDynamicFieldSelection()
    Call CreatePivotWithDynamicFieldSelection
End Sub

Sub CreatePivotWithDynamicFieldSelection()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim dataRange As Range
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim lastRow As Long
    Dim rowField As String
    Dim valueField As String
    
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("動態樞紐資料")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Worksheets.Add
        wsData.Name = "動態樞紐資料"
    End If
    
    wsData.Cells.Clear
    Call FillDynamicPivotData(wsData)
    
    rowField = InputBox("請輸入列欄位名稱（例如: 類別）:", "動態樞紐欄位選擇", "類別")
    
    If rowField = "" Then
        MsgBox "已取消操作。", vbInformation, "取消"
        Exit Sub
    End If
    
    valueField = InputBox("請輸入值欄位名稱（例如: 金額）:", "動態樞紐欄位選擇", "金額")
    
    If valueField = "" Then
        MsgBox "已取消操作。", vbInformation, "取消"
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("動態樞紐結果")
    If Not wsPivot Is Nothing Then wsPivot.Delete
    On Error GoTo 0
    
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "動態樞紐結果"
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set dataRange = wsData.Range("A1:C" & lastRow)
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange, _
        Version:=xlPivotTableVersion15)
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="DynamicPivot")
    
    On Error Resume Next
    pt.PivotFields(rowField).Orientation = xlRowField
    pt.PivotFields(rowField).Position = 1
    
    If Err.Number <> 0 Then
        MsgBox "欄位 '" & rowField & "' 不存在，請確認欄位名稱！", vbExclamation, "錯誤"
        wsPivot.Delete
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error Resume Next
    pt.AddDataField pt.PivotFields(valueField), valueField & "合計", xlSum
    
    If Err.Number <> 0 Then
        MsgBox "欄位 '" & valueField & "' 不存在，請確認欄位名稱！", vbExclamation, "錯誤"
        wsPivot.Delete
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    wsPivot.Activate
    
    MsgBox "動態樞紐分析表建立完成！" & vbCrLf & _
           "列欄位: " & rowField & vbCrLf & _
           "值欄位: " & valueField & vbCrLf & _
           "彙總方式: 加總", vbInformation, "完成"
End Sub

Private Sub FillDynamicPivotData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "類別"
    ws.Range("C1").Value = "金額"
    
    ws.Range("A2").Value = "2024/1/5"
    ws.Range("B2").Value = "食品"
    ws.Range("C2").Value = 1200
    ws.Range("A3").Value = "2024/1/10"
    ws.Range("B3").Value = "飲料"
    ws.Range("C3").Value = 800
    ws.Range("A4").Value = "2024/2/3"
    ws.Range("B4").Value = "食品"
    ws.Range("C4").Value = 1500
    ws.Range("A5").Value = "2024/2/15"
    ws.Range("B5").Value = "飲料"
    ws.Range("C5").Value = 650
    ws.Range("A6").Value = "2024/3/8"
    ws.Range("B6").Value = "日用品"
    ws.Range("C6").Value = 2200
    ws.Range("A7").Value = "2024/3/20"
    ws.Range("B7").Value = "食品"
    ws.Range("C7").Value = 1800
    ws.Range("A8").Value = "2024/4/2"
    ws.Range("B8").Value = "飲料"
    ws.Range("C8").Value = 950
    ws.Range("A9").Value = "2024/4/18"
    ws.Range("B9").Value = "日用品"
    ws.Range("C9").Value = 1600
    
    ws.Columns("A:C").AutoFit
End Sub
