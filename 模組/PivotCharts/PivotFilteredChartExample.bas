Attribute VB_Name = "PivotFilteredChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotFilteredChartExample
'功能說明: 建立已套用頁面篩選的樞紐分析圖，示範如何以 VBA 設定
'          樞紐圖的頁面欄位篩選條件，只顯示特定類別的資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

Sub CreatePivotFilteredChart()
    Dim ws      As Worksheet
    Dim wsPivot As Worksheet
    Dim pt      As PivotTable
    Dim pc      As PivotCache
    Dim chtObj  As ChartObject
    Dim oChart  As Chart

    ' 建立範例資料
    Set ws = GetOrCreateFilteredSheet("篩選圖表資料")
    Call FillFilteredChartData(ws)

    ' 移除舊工作表
    Dim wsOld As Worksheet
    On Error Resume Next
    Set wsOld = ThisWorkbook.Worksheets("篩選樞紐圖")
    On Error GoTo 0
    If Not wsOld Is Nothing Then
        Application.DisplayAlerts = False
        wsOld.Delete
        Application.DisplayAlerts = True
    End If

    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "篩選樞紐圖"

    ' 建立 PivotCache
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ws.Range("A1:D" & lastRow))

    ' 建立樞紐分析表
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="篩選樞紐")

    With pt
        ' 頁面篩選欄位（類別）
        .PivotFields("類別").Orientation = xlPageField
        .PivotFields("類別").Position = 1

        ' 列欄位（月份）
        .PivotFields("月份").Orientation = xlRowField
        .PivotFields("月份").Position = 1

        ' 欄欄位（業務員）
        .PivotFields("業務員").Orientation = xlColumnField
        .PivotFields("業務員").Position = 1

        ' 值欄位（銷售量）
        .PivotFields("銷售量").Orientation = xlDataField
        .PivotFields("銷售量").Function = xlSum
        .PivotFields("銷售量").NumberFormat = "#,##0"
    End With

    ' 套用頁面篩選：只顯示「電子產品」類別
    Call FilterPivotPageField(pt, "類別", "電子產品")

    ' 在樞紐旁邊建立樞紐圖
    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=300, Top:=10, Width:=450, Height:=280)
    Set oChart = chtObj.Chart

    oChart.SetSourceData Source:=pt.TableRange1
    oChart.ChartType = xlColumnClustered
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "電子產品月銷售量（已篩選）"
    oChart.HasLegend = True

    wsPivot.Columns.AutoFit
    MsgBox "篩選樞紐圖已建立！目前顯示：電子產品 類別。", vbInformation, "完成"
End Sub

' 設定頁面篩選欄位為指定值
Private Sub FilterPivotPageField(ByVal pt As PivotTable, _
                                  ByVal fieldName As String, _
                                  ByVal filterValue As String)
    Dim pf As PivotField
    Dim pi As PivotItem

    Set pf = pt.PivotFields(fieldName)
    pf.ClearAllFilters

    On Error Resume Next
    pf.CurrentPage = filterValue
    On Error GoTo 0
End Sub

Private Sub FillFilteredChartData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("月份", "業務員", "類別", "銷售量")

    Dim r As Long
    Dim data(1 To 12, 1 To 4) As Variant
    data(1, 1) = "一月":  data(1, 2) = "張志豪": data(1, 3) = "電子產品": data(1, 4) = 230
    data(2, 1) = "一月":  data(2, 2) = "李佳蓉": data(2, 3) = "電子產品": data(2, 4) = 180
    data(3, 1) = "一月":  data(3, 2) = "王大明": data(3, 3) = "家電":     data(3, 4) = 150
    data(4, 1) = "二月":  data(4, 2) = "張志豪": data(4, 3) = "電子產品": data(4, 4) = 270
    data(5, 1) = "二月":  data(5, 2) = "李佳蓉": data(5, 3) = "電子產品": data(5, 4) = 210
    data(6, 1) = "二月":  data(6, 2) = "王大明": data(6, 3) = "家電":     data(6, 4) = 130
    data(7, 1) = "三月":  data(7, 2) = "張志豪": data(7, 3) = "電子產品": data(7, 4) = 300
    data(8, 1) = "三月":  data(8, 2) = "李佳蓉": data(8, 3) = "電子產品": data(8, 4) = 250
    data(9, 1) = "三月":  data(9, 2) = "王大明": data(9, 3) = "家電":     data(9, 4) = 190
    data(10, 1) = "四月": data(10, 2) = "張志豪": data(10, 3) = "家電":   data(10, 4) = 120
    data(11, 1) = "四月": data(11, 2) = "李佳蓉": data(11, 3) = "電子產品": data(11, 4) = 220
    data(12, 1) = "四月": data(12, 2) = "王大明": data(12, 3) = "電子產品": data(12, 4) = 170

    For r = 1 To 12
        ws.Cells(r + 1, 1).Value = data(r, 1)
        ws.Cells(r + 1, 2).Value = data(r, 2)
        ws.Cells(r + 1, 3).Value = data(r, 3)
        ws.Cells(r + 1, 4).Value = data(r, 4)
    Next r
    ws.Columns("A:D").AutoFit
End Sub

Private Function GetOrCreateFilteredSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateFilteredSheet = ws
End Function
