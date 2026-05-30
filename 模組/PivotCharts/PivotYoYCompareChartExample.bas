Attribute VB_Name = "PivotYoYCompareChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotYoYCompareChartExample
'功能說明: 以VBA建立樞紐分析圖，比較同比（年度對比）銷售數據趨勢
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestPivotYoYCompareChart()
    Call CreateYoYPivotChart
End Sub

' 建立同比比較樞紐分析圖
Sub CreateYoYPivotChart()
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsChart As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim chObj As ChartObject
    Dim ch As Chart

    On Error GoTo ErrHandler
    Set wb = ThisWorkbook

    Set wsData = GetOrCreateYoYSheet(wb, "同比資料來源")
    Call FillYoYSalesData(wsData)

    Set wsChart = GetOrCreateYoYSheet(wb, "同比比較樞紐圖")

    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsChart.Range("A3"), _
        TableName:="同比樞紐")

    Set pf = pt.PivotFields("月份")
    pf.Orientation = xlRowField
    pf.Position = 1

    Set pf = pt.PivotFields("年度")
    pf.Orientation = xlColumnField
    pf.Position = 1

    Set pf = pt.PivotFields("銷售額")
    pf.Orientation = xlDataField
    pf.Function = xlSum
    pf.NumberFormat = "#,##0"
    pf.Name = "銷售額合計"

    Set chObj = wsChart.ChartObjects.Add( _
        Left:=wsChart.Range("I3").Left, _
        Top:=wsChart.Range("I3").Top, _
        Width:=480, _
        Height:=320)

    Set ch = chObj.Chart
    ch.SetSourceData Source:=pt.TableRange1
    ch.ChartType = xlLineMarkers

    ch.HasTitle = True
    ch.ChartTitle.Text = "年度銷售額同比比較"

    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom

    With ch.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "月份"
    End With

    With ch.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售額（元）"
    End With

    wsChart.Columns("A:H").AutoFit
    wsChart.Activate
    MsgBox "同比比較樞紐分析圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 填入同比資料（兩年度每月資料）
Private Sub FillYoYSalesData(ByVal ws As Worksheet)
    Dim months(1 To 12) As String
    Dim sales2024(1 To 12) As Long
    Dim sales2025(1 To 12) As Long
    Dim r As Integer
    Dim m As Integer

    ws.Cells.Clear
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "年度"
    ws.Range("C1").Value = "銷售額"
    ws.Range("A1:C1").Font.Bold = True

    months(1) = "1月" : months(2) = "2月" : months(3) = "3月"
    months(4) = "4月" : months(5) = "5月" : months(6) = "6月"
    months(7) = "7月" : months(8) = "8月" : months(9) = "9月"
    months(10) = "10月" : months(11) = "11月" : months(12) = "12月"

    sales2024(1) = 85000 : sales2024(2) = 72000 : sales2024(3) = 96000
    sales2024(4) = 110000 : sales2024(5) = 125000 : sales2024(6) = 118000
    sales2024(7) = 132000 : sales2024(8) = 140000 : sales2024(9) = 128000
    sales2024(10) = 115000 : sales2024(11) = 108000 : sales2024(12) = 155000

    sales2025(1) = 92000 : sales2025(2) = 78000 : sales2025(3) = 105000
    sales2025(4) = 118000 : sales2025(5) = 135000 : sales2025(6) = 122000
    sales2025(7) = 145000 : sales2025(8) = 152000 : sales2025(9) = 138000
    sales2025(10) = 125000 : sales2025(11) = 118000 : sales2025(12) = 168000

    r = 2
    For m = 1 To 12
        ws.Cells(r, 1).Value = months(m)
        ws.Cells(r, 2).Value = "2024年"
        ws.Cells(r, 3).Value = sales2024(m)
        r = r + 1
        ws.Cells(r, 1).Value = months(m)
        ws.Cells(r, 2).Value = "2025年"
        ws.Cells(r, 3).Value = sales2025(m)
        r = r + 1
    Next m
    ws.Columns.AutoFit
End Sub

' 取得或建立工作表並清除內容
Private Function GetOrCreateYoYSheet(ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateYoYSheet = ws
End Function
