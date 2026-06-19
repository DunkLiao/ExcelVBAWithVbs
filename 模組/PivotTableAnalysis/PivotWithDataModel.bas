Attribute VB_Name = "PivotWithDataModel"
Option Explicit
'*************************************************************************************
'模組名稱: PivotWithDataModel
'功能說明: 使用 Excel 資料模型建立樞紐分析表，將資料加入 Data Model 後進行分析
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestPivotWithDataModel()
    Call CreatePivotWithDataModel
End Sub

Sub CreatePivotWithDataModel()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim tbl As ListObject
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    ' 建立資料工作表
    Dim wsName As String
    wsName = "DataModel來源"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    ThisWorkbook.Sheets("DataModel樞紐").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set wsData = ThisWorkbook.Sheets.Add
    wsData.Name = wsName
    
    ' 撰寫範例資料
    wsData.Range("A1").Value = "訂單日期"
    wsData.Range("B1").Value = "產品類別"
    wsData.Range("C1").Value = "產品名稱"
    wsData.Range("D1").Value = "地區"
    wsData.Range("E1").Value = "數量"
    wsData.Range("F1").Value = "銷售額"
    
    Dim dataArr As Variant
    dataArr = Array( _
        Array(#1/5/2026#, "電子產品", "筆記型電腦", "北區", 3, 90000), _
        Array(#2/12/2026#, "電子產品", "平板電腦", "中區", 5, 75000), _
        Array(#3/8/2026#, "電子產品", "智慧型手機", "南區", 8, 176000), _
        Array(#4/15/2026#, "家居用品", "辦公椅", "北區", 12, 60000), _
        Array(#5/20/2026#, "家居用品", "書桌", "中區", 6, 90000), _
        Array(#6/3/2026#, "家居用品", "檯燈", "南區", 15, 45000), _
        Array(#7/18/2026#, "食品飲料", "咖啡豆", "北區", 20, 12000), _
        Array(#8/25/2026#, "食品飲料", "茶葉", "中區", 18, 9000), _
        Array(#9/10/2026#, "食品飲料", "礦泉水", "南區", 50, 25000), _
        Array(#10/30/2026#, "電子產品", "藍牙耳機", "北區", 25, 87500))
    
    Dim i As Long, j As Long
    For i = 0 To 9
        For j = 0 To 5
            wsData.Cells(i + 2, j + 1).Value = dataArr(i)(j)
        Next j
    Next i
    
    ' 將範圍轉換為表格（加入資料模型）
    Set tbl = wsData.ListObjects.Add(xlSrcRange, wsData.Range("A1:F11"), , xlYes)
    tbl.Name = "銷售明細"
    tbl.TableStyle = "TableStyleMedium6"
    
    ' 建立連結至資料模型的樞紐快取
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=tbl.Name, _
        Version:=6)
    
    ' 建立樞紐工作表
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "DataModel樞紐"
    
    ' 建立樞紐分析表
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="樞紐_資料模型", _
        DefaultVersion:=6)
    
    ' 設定樞紐欄位：產品類別為列標籤
    With pt
        .PivotFields("產品類別").Orientation = xlRowField
        .PivotFields("產品類別").Position = 1
        
        .PivotFields("地區").Orientation = xlColumnField
        .PivotFields("地區").Position = 1
        
        .PivotFields("銷售額").Orientation = xlDataField
        .PivotFields("銷售額").Function = xlSum
        .PivotFields("銷售額").NumberFormat = "#,##0"
        
        .ColumnGrand = True
        .RowGrand = True
    End With
    
    wsPivot.Columns.AutoFit
    wsData.Columns.AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "使用資料模型的樞紐分析表建立完成！" & vbCrLf & _
           "請查看「DataModel樞紐」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "建立樞紐分析表時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
