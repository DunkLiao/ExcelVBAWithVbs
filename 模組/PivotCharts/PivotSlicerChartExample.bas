Attribute VB_Name = "PivotSlicerChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotSlicerChartExample
'功能說明: 建立連結交叉分析篩選器（Slicer）的樞紐分析圖範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestPivotSlicerChart()
    Call CreatePivotSlicerChart
End Sub

' 建立帶有交叉分析篩選器的樞紐分析圖
Sub CreatePivotSlicerChart()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' 準備資料工作表
    Dim wsData As Worksheet
    Set wsData = GetOrCreateSlicerWs(wb, "篩選器圖資料")
    Call FillSlicerChartData(wsData)

    ' 準備樞紐工作表
    Dim wsPivot As Worksheet
    Set wsPivot = GetOrCreateSlicerWs(wb, "篩選器樞紐圖")

    ' 清除舊有樞紐
    Dim pt As PivotTable
    For Each pt In wsPivot.PivotTables
        pt.TableRange2.Clear
    Next pt

    ' 建立樞紐快取
    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.UsedRange)

    ' 建立樞紐分析表
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="SlicerPivot")

    With pt.PivotFields("月份")
        .Orientation = xlColumnField
        .Position = 1
    End With

    With pt.PivotFields("業績")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0"
        .Name = "業績合計"
    End With

    With pt.PivotFields("部門")
        .Orientation = xlRowField
        .Position = 1
    End With

    pt.TableStyle2 = "PivotStyleMedium2"

    ' 建立樞紐分析圖
    Dim chartObj As ChartObject
    Set chartObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("A20").Left, _
        Top:=wsPivot.Range("A20").Top, _
        Width:=480, Height:=280)

    Dim cht As Chart
    Set cht = chartObj.Chart
    cht.SetSourceData Source:=pt.TableRange2
    cht.ChartType = xlColumnClustered
    cht.HasTitle = True
    cht.ChartTitle.Text = "各部門月份業績（含篩選器）"

    ' 新增交叉分析篩選器（依部門）
    Dim sc As SlicerCache
    Set sc = wb.SlicerCaches.Add2(pt, "部門")
    sc.Slicers.Add wsPivot, , "部門篩選器", "部門", _
        wsPivot.Range("H3").Top, wsPivot.Range("H3").Left, 150, 200

    wsPivot.Activate
    MsgBox "帶有交叉分析篩選器的樞紐分析圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立樞紐分析圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 填入測試資料
Private Sub FillSlicerChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "月份"
    ws.Range("C1").Value = "業績"
    ws.Range("A1:C1").Font.Bold = True

    Dim rows(1 To 12, 1 To 3) As Variant
    rows(1, 1) = "業務部"  : rows(1, 2) = "1月"  : rows(1, 3) = 120000
    rows(2, 1) = "業務部"  : rows(2, 2) = "2月"  : rows(2, 3) = 135000
    rows(3, 1) = "業務部"  : rows(3, 2) = "3月"  : rows(3, 3) = 98000
    rows(4, 1) = "行銷部"  : rows(4, 2) = "1月"  : rows(4, 3) = 88000
    rows(5, 1) = "行銷部"  : rows(5, 2) = "2月"  : rows(5, 3) = 95000
    rows(6, 1) = "行銷部"  : rows(6, 2) = "3月"  : rows(6, 3) = 102000
    rows(7, 1) = "研發部"  : rows(7, 2) = "1月"  : rows(7, 3) = 65000
    rows(8, 1) = "研發部"  : rows(8, 2) = "2月"  : rows(8, 3) = 70000
    rows(9, 1) = "研發部"  : rows(9, 2) = "3月"  : rows(9, 3) = 68000
    rows(10, 1) = "客服部" : rows(10, 2) = "1月"  : rows(10, 3) = 45000
    rows(11, 1) = "客服部" : rows(11, 2) = "2月"  : rows(11, 3) = 52000
    rows(12, 1) = "客服部" : rows(12, 2) = "3月"  : rows(12, 3) = 49000

    Dim i As Long
    For i = 1 To 12
        ws.Cells(i + 1, 1).Value = rows(i, 1)
        ws.Cells(i + 1, 2).Value = rows(i, 2)
        ws.Cells(i + 1, 3).Value = rows(i, 3)
    Next i

    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSlicerWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSlicerWs = ws
End Function
