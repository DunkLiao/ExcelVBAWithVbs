Attribute VB_Name = "PageFieldPivot"
Option Explicit
'*************************************************************************************
'模組名稱: PageFieldPivot
'功能說明: 頁面欄位（報表篩選）操作示範
'          設定報表篩選欄位，並示範以 VBA 逐一切換篩選值
'          以及使用 ShowPages 依篩選值自動產生各頁工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestPageFieldPivot()
    Call CreatePageFieldPivot
End Sub

' 建立含頁面欄位的樞紐分析表，並示範切換篩選值
Sub CreatePageFieldPivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim pf As PivotField
    Dim pi As PivotItem

    Set wsData = GetOrCreateSheet(ThisWorkbook, "年度銷售")
    Call FillYearData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "頁面欄位樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="頁面欄位樞紐分析表")

    ' 頁面欄位（報表篩選）：年度
    With pt.PivotFields("年度")
        .Orientation = xlPageField
        .Position = 1
    End With

    ' 列欄位：地區
    With pt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 欄欄位：產品
    With pt.PivotFields("產品")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' 值欄位：銷售額
    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum

    ' 示範以 VBA 切換頁面篩選值為「2025」
    Set pf = pt.PivotFields("年度")
    pf.CurrentPage = "2025"

    wsPivot.Columns("A:F").AutoFit
    wsPivot.Activate

    ' 提示：ShowPages 示範（會產生新工作表，已備註）
    ' pt.ShowPages FieldName:="年度"

    MsgBox "頁面欄位樞紐分析表已建立！" & Chr(13) & _
           "目前篩選年度：2025" & Chr(13) & _
           "可修改 CurrentPage 屬性切換其他年度。", vbInformation, "完成"
End Sub

' 填入年度銷售資料
Private Sub FillYearData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "年度"
    ws.Range("B1").Value = "地區"
    ws.Range("C1").Value = "產品"
    ws.Range("D1").Value = "銷售額"

    ws.Range("A2").Value = "2024"  : ws.Range("B2").Value = "北部" : ws.Range("C2").Value = "A產品" : ws.Range("D2").Value = 80000
    ws.Range("A3").Value = "2024"  : ws.Range("B3").Value = "南部" : ws.Range("C3").Value = "A產品" : ws.Range("D3").Value = 65000
    ws.Range("A4").Value = "2024"  : ws.Range("B4").Value = "北部" : ws.Range("C4").Value = "B產品" : ws.Range("D4").Value = 72000
    ws.Range("A5").Value = "2024"  : ws.Range("B5").Value = "南部" : ws.Range("C5").Value = "B產品" : ws.Range("D5").Value = 55000
    ws.Range("A6").Value = "2025"  : ws.Range("B6").Value = "北部" : ws.Range("C6").Value = "A產品" : ws.Range("D6").Value = 95000
    ws.Range("A7").Value = "2025"  : ws.Range("B7").Value = "南部" : ws.Range("C7").Value = "A產品" : ws.Range("D7").Value = 78000
    ws.Range("A8").Value = "2025"  : ws.Range("B8").Value = "北部" : ws.Range("C8").Value = "B產品" : ws.Range("D8").Value = 88000
    ws.Range("A9").Value = "2025"  : ws.Range("B9").Value = "南部" : ws.Range("C9").Value = "B產品" : ws.Range("D9").Value = 69000

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
