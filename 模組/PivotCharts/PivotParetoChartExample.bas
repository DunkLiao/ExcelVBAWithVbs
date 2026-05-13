Attribute VB_Name = "PivotParetoChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotParetoChartExample
'功能說明: 以樞紐分析表為資料來源，建立柏拉圖（Pareto）組合圖範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub CreatePivotParetoChart()
    Dim wb          As Workbook
    Dim dataWs      As Worksheet
    Dim pivotWs     As Worksheet
    Dim pc          As PivotCache
    Dim pt          As PivotTable
    Dim chartObj    As ChartObject
    Dim cht         As Chart
    Dim dataRange   As Range

    Set wb = ThisWorkbook

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("柏拉圖資料").Delete
    wb.Worksheets("柏拉圖樞紐").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set dataWs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    dataWs.Name = "柏拉圖資料"

    dataWs.Range("A1").Value = "缺陷類型"
    dataWs.Range("B1").Value = "數量"
    dataWs.Range("A2").Value = "外觀刮傷" : dataWs.Range("B2").Value = 85
    dataWs.Range("A3").Value = "尺寸偏差" : dataWs.Range("B3").Value = 60
    dataWs.Range("A4").Value = "色差問題" : dataWs.Range("B4").Value = 42
    dataWs.Range("A5").Value = "功能異常" : dataWs.Range("B5").Value = 30
    dataWs.Range("A6").Value = "包裝破損" : dataWs.Range("B6").Value = 18
    dataWs.Range("A7").Value = "其他"     : dataWs.Range("B7").Value = 10
    dataWs.Range("A1:B1").Font.Bold = True

    Set pivotWs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    pivotWs.Name = "柏拉圖樞紐"

    Set dataRange = dataWs.Range("A1:B7")
    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="柏拉圖PT")

    With pt
        .PivotFields("缺陷類型").Orientation = xlRowField
        .PivotFields("缺陷類型").Position = 1
        .AddDataField .PivotFields("數量"), "加總-數量", xlSum
    End With

    pt.TableStyle2 = "PivotStyleMedium9"

    Set chartObj = pivotWs.ChartObjects.Add( _
        Left:=pivotWs.Range("E1").Left, _
        Top:=pivotWs.Range("E1").Top, _
        Width:=480, Height:=320)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=pt.TableRange2
    cht.ChartType = xlColumnClustered

    cht.HasTitle = True
    cht.ChartTitle.Text = "缺陷類型柏拉圖分析"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "缺陷類型"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "缺陷數量"
    End With

    cht.ChartStyle = 10
    cht.HasLegend = True

    MsgBox "樞紐柏拉圖已建立完成！", vbInformation, "完成"
End Sub
