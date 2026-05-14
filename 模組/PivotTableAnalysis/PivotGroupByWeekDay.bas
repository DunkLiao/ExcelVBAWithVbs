Attribute VB_Name = "PivotGroupByWeekDay"
Option Explicit
'*************************************************************************************
'模組名稱: PivotGroupByWeekDay
'功能說明: 建立樞紐分析表，依星期（一～日）分組並統計業績資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestPivotGroupByWeekDay()
    Call CreateWeekDayPivot
End Sub

' 建立依星期分組的樞紐分析表
Sub CreateWeekDayPivot()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim wsData As Worksheet
    Set wsData = GetOrCreateWdWs(wb, "星期業績資料")
    Call FillWeekDayData(wsData)

    Dim wsPivot As Worksheet
    Set wsPivot = GetOrCreateWdWs(wb, "星期樞紐分析")

    ' 刪除舊有樞紐快取
    Dim pc As PivotCache
    Dim pt As PivotTable
    For Each pt In wsPivot.PivotTables
        pt.TableRange2.Clear
    Next pt

    ' 建立樞紐快取
    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.UsedRange)

    ' 建立樞紐分析表
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="WeekDayPivot")

    ' 設定列欄位（星期名稱）
    With pt.PivotFields("星期")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 設定值欄位（業績）
    With pt.PivotFields("業績")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0"
        .Name = "業績合計"
    End With

    ' 設定計數欄位
    With pt.PivotFields("業績")
        .Orientation = xlDataField
        .Function = xlCount
        .Name = "筆數"
    End With

    ' 依星期排序：週一→週日
    Dim weekOrder(0 To 6) As String
    weekOrder(0) = "週一"
    weekOrder(1) = "週二"
    weekOrder(2) = "週三"
    weekOrder(3) = "週四"
    weekOrder(4) = "週五"
    weekOrder(5) = "週六"
    weekOrder(6) = "週日"

    Dim pf As PivotField
    Set pf = pt.PivotFields("星期")
    Dim idx As Long
    For idx = 0 To 6
        On Error Resume Next
        pf.PivotItems(weekOrder(idx)).Position = idx + 1
        On Error GoTo 0
    Next idx

    pt.TableStyle2 = "PivotStyleMedium9"
    wsPivot.Activate

    MsgBox "依星期分組的樞紐分析表已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立樞紐分析表時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 填入含星期名稱的測試資料
Private Sub FillWeekDayData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "星期"
    ws.Range("C1").Value = "業務員"
    ws.Range("D1").Value = "業績"
    ws.Range("A1:D1").Font.Bold = True

    Dim sampleData(1 To 14, 1 To 4) As Variant
    sampleData(1, 1) = "2026/5/4"  : sampleData(1, 2) = "週一" : sampleData(1, 3) = "王小明" : sampleData(1, 4) = 12000
    sampleData(2, 1) = "2026/5/5"  : sampleData(2, 2) = "週二" : sampleData(2, 3) = "李大華" : sampleData(2, 4) = 9500
    sampleData(3, 1) = "2026/5/6"  : sampleData(3, 2) = "週三" : sampleData(3, 3) = "陳美玲" : sampleData(3, 4) = 15000
    sampleData(4, 1) = "2026/5/7"  : sampleData(4, 2) = "週四" : sampleData(4, 3) = "林俊傑" : sampleData(4, 4) = 8000
    sampleData(5, 1) = "2026/5/8"  : sampleData(5, 2) = "週五" : sampleData(5, 3) = "張志遠" : sampleData(5, 4) = 18000
    sampleData(6, 1) = "2026/5/9"  : sampleData(6, 2) = "週六" : sampleData(6, 3) = "王小明" : sampleData(6, 4) = 6000
    sampleData(7, 1) = "2026/5/10" : sampleData(7, 2) = "週日" : sampleData(7, 3) = "李大華" : sampleData(7, 4) = 4000
    sampleData(8, 1) = "2026/5/11" : sampleData(8, 2) = "週一" : sampleData(8, 3) = "陳美玲" : sampleData(8, 4) = 13500
    sampleData(9, 1) = "2026/5/12" : sampleData(9, 2) = "週二" : sampleData(9, 3) = "林俊傑" : sampleData(9, 4) = 11000
    sampleData(10, 1) = "2026/5/13" : sampleData(10, 2) = "週三" : sampleData(10, 3) = "張志遠" : sampleData(10, 4) = 16500
    sampleData(11, 1) = "2026/5/14" : sampleData(11, 2) = "週四" : sampleData(11, 3) = "王小明" : sampleData(11, 4) = 9000
    sampleData(12, 1) = "2026/5/15" : sampleData(12, 2) = "週五" : sampleData(12, 3) = "李大華" : sampleData(12, 4) = 21000
    sampleData(13, 1) = "2026/5/16" : sampleData(13, 2) = "週六" : sampleData(13, 3) = "陳美玲" : sampleData(13, 4) = 7500
    sampleData(14, 1) = "2026/5/17" : sampleData(14, 2) = "週日" : sampleData(14, 3) = "林俊傑" : sampleData(14, 4) = 3500

    Dim i As Long
    For i = 1 To 14
        ws.Cells(i + 1, 1).Value = sampleData(i, 1)
        ws.Cells(i + 1, 2).Value = sampleData(i, 2)
        ws.Cells(i + 1, 3).Value = sampleData(i, 3)
        ws.Cells(i + 1, 4).Value = sampleData(i, 4)
    Next i

    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateWdWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateWdWs = ws
End Function
