Attribute VB_Name = "DisplayModePivot"
Option Explicit
'*************************************************************************************
'模組名稱: DisplayModePivot
'功能說明: 樞紐顯示模式切換示範
'          示範三種版面配置模式：
'          xlCompactRow（精簡）、xlOutlineRow（大綱）、xlTabularRow（表格）
'          以及小計與總計的顯示控制
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestDisplayModePivot()
    Call CreateDisplayModePivot
End Sub

' 建立樞紐分析表並示範三種版面配置模式
Sub CreateDisplayModePivot()
    Dim wsData As Worksheet
    Dim pc As PivotCache
    Dim dataRange As Range

    Set wsData = GetOrCreateSheet(ThisWorkbook, "組織銷售")
    Call FillOrgData(wsData)

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' 建立三個工作表，各套用一種顯示模式
    Call BuildModeSheet(pc, dataRange, "精簡模式樞紐", "精簡模式樞紐分析表", xlCompactRow)
    Call BuildModeSheet(pc, dataRange, "大綱模式樞紐", "大綱模式樞紐分析表", xlOutlineRow)
    Call BuildModeSheet(pc, dataRange, "表格模式樞紐", "表格模式樞紐分析表", xlTabularRow)

    ThisWorkbook.Worksheets("精簡模式樞紐").Activate
    MsgBox "三種版面配置模式示範完成！" & Chr(13) & _
           "‧ 精簡模式樞紐：多欄位合併顯示在同一欄" & Chr(13) & _
           "‧ 大綱模式樞紐：多欄位依層級縮排顯示" & Chr(13) & _
           "‧ 表格模式樞紐：多欄位各自獨立欄顯示", vbInformation, "完成"
End Sub

' 依指定版面模式建立單一樞紐分析表工作表
Private Sub BuildModeSheet( _
    ByVal pc As PivotCache, _
    ByVal dataRange As Range, _
    ByVal sheetName As String, _
    ByVal tableName As String, _
    ByVal layoutMode As Long)

    Dim wsPivot As Worksheet
    Dim pt As PivotTable

    Set wsPivot = GetOrCreateSheet(ThisWorkbook, sheetName)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:=tableName)

    ' 多層列欄位：部門 → 組別
    With pt.PivotFields("部門")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields("組別")
        .Orientation = xlRowField
        .Position = 2
    End With

    ' 欄欄位：季度
    With pt.PivotFields("季度")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' 值欄位：業績
    Dim df As PivotField
    Set df = pt.AddDataField(pt.PivotFields("業績"), "業績合計", xlSum)
    df.NumberFormat = "#,##0"

    ' 套用版面配置模式
    pt.RowAxisLayout layoutMode

    ' 控制小計顯示（大綱/表格模式下可見）
    pt.PivotFields("部門").Subtotals(1) = True  ' 顯示自動小計

    ' 控制總計顯示
    pt.ColumnGrand = True
    pt.RowGrand = True

    wsPivot.Columns("A:G").AutoFit
End Sub

' 填入組織業績資料
Private Sub FillOrgData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "組別"
    ws.Range("C1").Value = "季度"
    ws.Range("D1").Value = "業績"

    ws.Range("A2").Value = "業務部" : ws.Range("B2").Value = "一組" : ws.Range("C2").Value = "Q1" : ws.Range("D2").Value = 250000
    ws.Range("A3").Value = "業務部" : ws.Range("B3").Value = "一組" : ws.Range("C3").Value = "Q2" : ws.Range("D3").Value = 310000
    ws.Range("A4").Value = "業務部" : ws.Range("B4").Value = "二組" : ws.Range("C4").Value = "Q1" : ws.Range("D4").Value = 190000
    ws.Range("A5").Value = "業務部" : ws.Range("B5").Value = "二組" : ws.Range("C5").Value = "Q2" : ws.Range("D5").Value = 225000
    ws.Range("A6").Value = "研發部" : ws.Range("B6").Value = "前端組" : ws.Range("C6").Value = "Q1" : ws.Range("D6").Value = 120000
    ws.Range("A7").Value = "研發部" : ws.Range("B7").Value = "前端組" : ws.Range("C7").Value = "Q2" : ws.Range("D7").Value = 145000
    ws.Range("A8").Value = "研發部" : ws.Range("B8").Value = "後端組" : ws.Range("C8").Value = "Q1" : ws.Range("D8").Value = 160000
    ws.Range("A9").Value = "研發部" : ws.Range("B9").Value = "後端組" : ws.Range("C9").Value = "Q2" : ws.Range("D9").Value = 175000

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
