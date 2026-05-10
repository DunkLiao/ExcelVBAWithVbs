Attribute VB_Name = "ShowValueAsPercentPivot"
Option Explicit

' ============================================================
' 模組名稱：ShowValueAsPercentPivot
' 功能說明：建立樞紐分析表並將值欄位設定為佔欄總計百分比
'           同時保留原始數量欄位供對照
' 使用方式：確認資料工作表名稱後執行，依提示操作
' ============================================================

Sub ShowValueAsPercentPivot()
    Dim wb          As Workbook
    Dim wsSrc       As Worksheet
    Dim wsPvt       As Worksheet
    Dim pvtCache    As PivotCache
    Dim pvt         As PivotTable
    Dim pvtName     As String
    Dim pvtShName   As String
    Dim srcRange    As Range
    Dim lastRow     As Long
    Dim lastCol     As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Set wb = ThisWorkbook
    pvtShName = "百分比樞紐"
    pvtName = "PvtPercent"
    
    ' 若無資料，建立範例資料
    Dim wsData As Worksheet
    Dim dataShName As String
    dataShName = "銷售資料"
    
    On Error Resume Next
    Set wsData = wb.Sheets(dataShName)
    On Error GoTo ErrHandler
    
    If wsData Is Nothing Then
        Call CreateSampleData(wb, dataShName)
        Set wsData = wb.Sheets(dataShName)
    End If
    
    ' 計算資料範圍
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set srcRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))
    
    ' 刪除舊的樞紐工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(pvtShName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    
    ' 新增樞紐工作表
    Set wsPvt = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsPvt.Name = pvtShName
    
    ' 建立樞紐快取
    Set pvtCache = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcRange)
    
    ' 建立樞紐分析表
    Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=wsPvt.Range("B2"), _
        TableName:=pvtName)
    
    ' 設定列欄位（地區）
    With pvt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ' 設定欄欄位（產品）
    With pvt.PivotFields("產品")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    ' 加入數量欄位（原始值）
    With pvt.PivotFields("銷售數量")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "數量合計"
        .NumberFormat = "#,##0"
    End With
    
    ' 加入數量欄位（佔欄百分比）
    With pvt.PivotFields("銷售數量")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "佔欄%"
        .Calculation = xlPercentOfColumn
        .NumberFormat = "0.0%"
    End With
    
    ' 套用樞紐樣式
    pvt.TableStyle2 = "PivotStyleMedium9"
    
    wsPvt.Columns.AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "百分比樞紐分析表建立完成！" & vbCrLf & _
           "請查看「" & pvtShName & "」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

' 建立範例銷售資料
Private Sub CreateSampleData(ByVal wb As Workbook, ByVal shName As String)
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = shName
    
    ws.Range("A1:C1").Value = Array("地區", "產品", "銷售數量")
    ws.Rows(1).Font.Bold = True
    
    Dim data(14, 3) As Variant
    Dim i As Integer
    
    ' 填入地區/產品/數量範例資料
    Dim regions(2) As String
    Dim products(2) As String
    regions(0) = "北區" : regions(1) = "中區" : regions(2) = "南區"
    products(0) = "產品A" : products(1) = "產品B" : products(2) = "產品C"
    
    Dim row As Integer
    row = 2
    Dim r As Integer, p As Integer
    For r = 0 To 2
        For p = 0 To 2
            ws.Cells(row, 1).Value = regions(r)
            ws.Cells(row, 2).Value = products(p)
            ws.Cells(row, 3).Value = Int(Rnd * 500) + 100
            row = row + 1
        Next p
    Next r
    
    ws.Columns.AutoFit
End Sub