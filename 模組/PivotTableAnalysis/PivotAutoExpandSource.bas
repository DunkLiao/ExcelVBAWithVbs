Attribute VB_Name = "PivotAutoExpandSource"
Option Explicit
'*************************************************************************************
'模組名稱: PivotAutoExpandSource
'功能說明: 建立可自動擴展資料來源範圍的樞紐分析表（使用動態範圍或表格）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestPivotAutoExpandSource()
    Call CreatePivotWithDynamicSource
End Sub

Sub CreatePivotWithDynamicSource()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim lo As ListObject
    Dim pc As PivotCache
    Dim pt As PivotTable

    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsData = ThisWorkbook.Worksheets("動態來源資料")
    If Not wsData Is Nothing Then wsData.Delete
    Set wsPivot = ThisWorkbook.Worksheets("動態來源樞紐")
    If Not wsPivot Is Nothing Then wsPivot.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 建立資料工作表
    Set wsData = ThisWorkbook.Worksheets.Add
    wsData.Name = "動態來源資料"

    wsData.Range("A1").Value = "日期"
    wsData.Range("B1").Value = "類別"
    wsData.Range("C1").Value = "金額"
    wsData.Range("A1:C1").Font.Bold = True

    wsData.Range("A2").Value = "2026/1/5"
    wsData.Range("B2").Value = "食品"
    wsData.Range("C2").Value = 2500

    wsData.Range("A3").Value = "2026/1/12"
    wsData.Range("B3").Value = "飲料"
    wsData.Range("C3").Value = 1800

    wsData.Range("A4").Value = "2026/1/19"
    wsData.Range("B4").Value = "食品"
    wsData.Range("C4").Value = 3200

    wsData.Range("A5").Value = "2026/1/26"
    wsData.Range("B5").Value = "日用品"
    wsData.Range("C5").Value = 1500

    ' 將資料轉換為 ListObject 表格（可自動擴展）
    Set lo = wsData.ListObjects.Add( _
        xlSrcRange, wsData.Range("A1").CurrentRegion, , xlYes)
    lo.Name = "銷售資料表"

    wsData.Columns("A:C").AutoFit

    ' 以表格為資料來源建立樞紐分析表
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "動態來源樞紐"

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=lo.Name)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="動態樞紐分析")

    With pt
        .AddFields RowFields:="類別"
        With .PivotFields("金額")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "合計金額"
        End With
    End With

    MsgBox "動態來源樞紐分析表已建立！" & vbCrLf & vbCrLf & _
           "新增資料到「動態來源資料」後，" & vbCrLf & _
           "重新整理樞紐即可自動納入新資料。", vbInformation, "完成"
End Sub
