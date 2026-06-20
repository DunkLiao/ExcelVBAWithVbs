Attribute VB_Name = "PivotCacheManagement"
Option Explicit
'*************************************************************************************
'模組名稱: PivotCacheManagement
'功能說明: 示範樞紐分析表快取管理，包括共用快取、快取資訊查詢與最佳化的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestPivotCacheManagement()
    Call PivotCacheManagementDemo
End Sub

' 樞紐快取管理示範
Sub PivotCacheManagementDemo()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim wsPivot2 As Worksheet
    Dim wsInfo As Worksheet
    Dim dataRange As Range
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pt2 As PivotTable
    Dim lastRow As Long
    Dim i As Long
    
    ' 建立資料工作表
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("快取管理資料")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Worksheets.Add
        wsData.Name = "快取管理資料"
    End If
    
    wsData.Cells.Clear
    Call FillCacheDemoData(wsData)
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set dataRange = wsData.Range("A1:C" & lastRow)
    
    ' 建立第一個樞紐分析表（使用新快取）
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("樞紐共用快取1")
    If Not wsPivot Is Nothing Then wsPivot.Delete
    On Error GoTo 0
    
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "樞紐共用快取1"
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange, _
        Version:=xlPivotTableVersion15)
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="PivotShared1")
    
    With pt
        .PivotFields("分類").Orientation = xlRowField
        .PivotFields("分類").Position = 1
        .PivotFields("月份").Orientation = xlColumnField
        .PivotFields("月份").Position = 1
        .PivotFields("金額").Orientation = xlDataField
    End With
    
    ' 建立第二個樞紐分析表（共用相同快取）
    On Error Resume Next
    Set wsPivot2 = ThisWorkbook.Worksheets("樞紐共用快取2")
    If Not wsPivot2 Is Nothing Then wsPivot2.Delete
    On Error GoTo 0
    
    Set wsPivot2 = ThisWorkbook.Worksheets.Add
    wsPivot2.Name = "樞紐共用快取2"
    
    Set pt2 = pc.CreatePivotTable( _
        TableDestination:=wsPivot2.Range("A1"), _
        TableName:="PivotShared2")
    
    With pt2
        .PivotFields("分類").Orientation = xlRowField
        .PivotFields("分類").Position = 1
    End With
    
    pt2.AddDataField pt2.PivotFields("金額"), "金額合計", xlSum
    
    ' 建立快取資訊工作表
    On Error Resume Next
    Set wsInfo = ThisWorkbook.Worksheets("快取資訊")
    If Not wsInfo Is Nothing Then wsInfo.Delete
    On Error GoTo 0
    
    Set wsInfo = ThisWorkbook.Worksheets.Add
    wsInfo.Name = "快取資訊"
    
    wsInfo.Range("A1").Value = "樞紐快取管理資訊"
    wsInfo.Range("A1").Font.Bold = True
    wsInfo.Range("A2").Value = "樞紐快取總數："
    wsInfo.Range("B2").Value = ThisWorkbook.PivotCaches.Count
    
    wsInfo.Range("A4").Value = "快取索引"
    wsInfo.Range("B4").Value = "資料筆數"
    
    For i = 1 To ThisWorkbook.PivotCaches.Count
        wsInfo.Cells(4 + i, 1).Value = ThisWorkbook.PivotCaches(i).Index
        
        On Error Resume Next
        wsInfo.Cells(4 + i, 2).Value = ThisWorkbook.PivotCaches(i).RecordCount
        On Error GoTo 0
    Next i
    
    Dim infoRow As Long
    infoRow = 4 + ThisWorkbook.PivotCaches.Count + 1
    wsInfo.Cells(infoRow, 1).Value = "說明：以上兩個樞紐分析表共用同一個快取，可節省記憶體。"
    
    wsInfo.Columns("A:C").AutoFit
    wsInfo.Activate
    
    MsgBox "樞紐快取管理示範完成！" & vbCrLf & _
           "請查看「快取資訊」工作表了解快取使用狀況。", vbInformation, "完成"
End Sub

' 填入快取管理示範資料
Private Sub FillCacheDemoData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "分類"
    ws.Range("B1").Value = "月份"
    ws.Range("C1").Value = "金額"
    
    ws.Range("A2").Value = "食品"
    ws.Range("B2").Value = "1月"
    ws.Range("C2").Value = 5000
    
    ws.Range("A3").Value = "飲料"
    ws.Range("B3").Value = "1月"
    ws.Range("C3").Value = 3200
    
    ws.Range("A4").Value = "食品"
    ws.Range("B4").Value = "2月"
    ws.Range("C4").Value = 4800
    
    ws.Range("A5").Value = "飲料"
    ws.Range("B5").Value = "2月"
    ws.Range("C5").Value = 3500
    
    ws.Range("A6").Value = "食品"
    ws.Range("B6").Value = "3月"
    ws.Range("C6").Value = 6200
    
    ws.Range("A7").Value = "飲料"
    ws.Range("B7").Value = "3月"
    ws.Range("C7").Value = 4100
    
    ws.Range("A8").Value = "百貨"
    ws.Range("B8").Value = "1月"
    ws.Range("C8").Value = 7200
    
    ws.Range("A9").Value = "百貨"
    ws.Range("B9").Value = "2月"
    ws.Range("C9").Value = 6800
    
    ws.Range("A10").Value = "百貨"
    ws.Range("B10").Value = "3月"
    ws.Range("C10").Value = 8100
    
    ws.Columns("A:C").AutoFit
End Sub
