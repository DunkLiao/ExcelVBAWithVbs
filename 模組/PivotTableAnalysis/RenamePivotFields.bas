Attribute VB_Name = "RenamePivotFields"
Option Explicit
'*************************************************************************************
'模組名稱: RenamePivotFields
'功能說明: 示範如何將樞紐分析表中的欄位標題重新命名為更易讀的顯示名稱
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口
Sub TestRenamePivotFields()
    Call CreatePivotWithRenamedFields("樞紐重命名範例")
End Sub

' 建立樞紐分析表並重新命名欄位
' sheetName: 工作表名稱
Sub CreatePivotWithRenamedFields(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache
    Dim pf As PivotField
    Dim dataRange As Range
    Dim pivotSheetName As String

    pivotSheetName = sheetName & "_樞紐"

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Call FillPivotRenameData(ws)
    Set dataRange = ws.Range("A1:D11")

    On Error Resume Next
    Set pivotWs = ThisWorkbook.Worksheets(pivotSheetName)
    On Error GoTo 0

    If pivotWs Is Nothing Then
        Set pivotWs = ThisWorkbook.Worksheets.Add(After:=ws)
        pivotWs.Name = pivotSheetName
    Else
        pivotWs.Cells.Clear
    End If

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=pivotWs.Range("B2"), _
        TableName:="PivotRenamed")

    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Product").Orientation = xlColumnField
        Set pf = .PivotFields("Sales")
        pf.Orientation = xlDataField
        pf.Function = xlSum
        pf.Name = "銷售合計"
    End With

    ' 重新命名列/欄欄位的顯示標題
    On Error Resume Next
    pt.PivotFields("Region").Caption = "銷售地區"
    pt.PivotFields("Product").Caption = "產品類別"
    On Error GoTo 0

    pivotWs.Columns.AutoFit
    MsgBox "樞紐分析表已建立並完成欄位重新命名！", vbInformation, "完成"
End Sub

' 填入樞紐範例資料
Private Sub FillPivotRenameData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "Region"
    ws.Range("B1").Value = "Product"
    ws.Range("C1").Value = "Month"
    ws.Range("D1").Value = "Sales"

    ws.Range("A2").Value = "北部":  ws.Range("B2").Value = "手機": ws.Range("C2").Value = "一月": ws.Range("D2").Value = 12000
    ws.Range("A3").Value = "北部":  ws.Range("B3").Value = "電腦": ws.Range("C3").Value = "一月": ws.Range("D3").Value = 28000
    ws.Range("A4").Value = "南部":  ws.Range("B4").Value = "手機": ws.Range("C4").Value = "一月": ws.Range("D4").Value = 9500
    ws.Range("A5").Value = "南部":  ws.Range("B5").Value = "電腦": ws.Range("C5").Value = "一月": ws.Range("D5").Value = 22000
    ws.Range("A6").Value = "東部":  ws.Range("B6").Value = "手機": ws.Range("C6").Value = "一月": ws.Range("D6").Value = 7800
    ws.Range("A7").Value = "北部":  ws.Range("B7").Value = "手機": ws.Range("C7").Value = "二月": ws.Range("D7").Value = 13500
    ws.Range("A8").Value = "北部":  ws.Range("B8").Value = "電腦": ws.Range("C8").Value = "二月": ws.Range("D8").Value = 31000
    ws.Range("A9").Value = "南部":  ws.Range("B9").Value = "手機": ws.Range("C9").Value = "二月": ws.Range("D9").Value = 10200
    ws.Range("A10").Value = "南部": ws.Range("B10").Value = "電腦": ws.Range("C10").Value = "二月": ws.Range("D10").Value = 24500
    ws.Range("A11").Value = "東部": ws.Range("B11").Value = "手機": ws.Range("C11").Value = "二月": ws.Range("D11").Value = 8300
    ws.Columns("A:D").AutoFit
End Sub