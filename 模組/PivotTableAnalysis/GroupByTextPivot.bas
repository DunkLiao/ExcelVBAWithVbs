Attribute VB_Name = "GroupByTextPivot"
Option Explicit
'*************************************************************************************
'模組名稱: GroupByTextPivot
'功能說明: 手動文字群組示範
'          將樞紐分析表中多個列項目合併為自訂群組
'          例如：台北+基隆 → 北部地區、高雄+台南 → 南部地區
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestGroupByTextPivot()
    Call CreateGroupByTextPivot
End Sub

' 建立含手動文字群組的樞紐分析表
Sub CreateGroupByTextPivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim dataRange As Range

    Set wsData = GetOrCreateSheet(ThisWorkbook, "城市銷售")
    Call FillCityData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "文字群組樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="文字群組樞紐分析表")

    ' 列欄位：城市
    With pt.PivotFields("城市")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 值欄位：銷售額
    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum

    ' 手動群組：台北+基隆 → 北部地區
    Set pf = pt.PivotFields("城市")
    pf.PivotItems("台北").Selected = True
    pf.PivotItems("基隆").Selected = True

    On Error Resume Next
    pf.CreatePivotItemList(Array("台北", "基隆")).Group
    pt.PivotFields("城市2").PivotItems("群組1").Caption = "北部地區"

    ' 手動群組：高雄+台南 → 南部地區
    pf.PivotItems("高雄").Selected = True
    pf.PivotItems("台南").Selected = True
    pf.CreatePivotItemList(Array("高雄", "台南")).Group
    pt.PivotFields("城市2").PivotItems("群組2").Caption = "南部地區"
    On Error GoTo 0

    wsPivot.Columns("A:D").AutoFit
    wsPivot.Activate
    MsgBox "手動文字群組樞紐分析表已建立！" & Chr(13) & _
           "城市已依地區分組顯示。", vbInformation, "完成"
End Sub

' 填入城市銷售資料
Private Sub FillCityData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "城市"
    ws.Range("B1").Value = "類別"
    ws.Range("C1").Value = "銷售額"

    ws.Range("A2").Value = "台北" : ws.Range("B2").Value = "電子" : ws.Range("C2").Value = 12000
    ws.Range("A3").Value = "基隆" : ws.Range("B3").Value = "服飾" : ws.Range("C3").Value = 8000
    ws.Range("A4").Value = "高雄" : ws.Range("B4").Value = "電子" : ws.Range("C4").Value = 15000
    ws.Range("A5").Value = "台南" : ws.Range("B5").Value = "服飾" : ws.Range("C5").Value = 9500
    ws.Range("A6").Value = "台北" : ws.Range("B6").Value = "食品" : ws.Range("C6").Value = 7000
    ws.Range("A7").Value = "高雄" : ws.Range("B7").Value = "食品" : ws.Range("C7").Value = 11000
    ws.Range("A8").Value = "基隆" : ws.Range("B8").Value = "電子" : ws.Range("C8").Value = 6500
    ws.Range("A9").Value = "台南" : ws.Range("B9").Value = "電子" : ws.Range("C9").Value = 13000

    ws.Columns("A:C").AutoFit
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
