Attribute VB_Name = "PivotStockChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotStockChartExample
'功能說明: 以樞紐分析表為資料來源建立股價走勢折線圖範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub CreatePivotStockChart()
    On Error GoTo ErrHandler
    Dim wb       As Workbook
    Dim wsData   As Worksheet
    Dim wsPivot  As Worksheet
    Dim pc       As PivotCache
    Dim pt       As PivotTable
    Dim chartObj As ChartObject
    Dim cht      As Chart
    Dim ptRange  As Range

    Set wb = ThisWorkbook
    Set wsData  = GetOrCreateSheetStk(wb, "股價來源資料")
    Call FillStockSourceData(wsData)
    Set wsPivot = GetOrCreateSheetStk(wb, "股價樞紐圖")

    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="股價PT")

    With pt.PivotFields("日期")
        .Orientation = xlRowField
        .Position     = 1
    End With

    With pt.PivotFields("收盤價")
        .Orientation  = xlDataField
        .Function     = xlAverage
        .NumberFormat = "#,##0.00"
        .Name          = "平均收盤價"
    End With

    pt.TableStyle2 = "PivotStyleMedium2"
    wsPivot.Columns("A:B").AutoFit
    Set ptRange = pt.TableRange1

    Set chartObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("D3").Left, _
        Top:=wsPivot.Range("D3").Top, _
        Width:=450, _
        Height:=280)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=ptRange
    cht.ChartType = xlLine
    cht.HasTitle = True
    cht.ChartTitle.Text = "每日平均收盤價走勢"
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "交易日"
    End With
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "收盤價（元）"
    End With
    cht.ChartStyle = 4
    wsPivot.Activate
    MsgBox "股價樞紐走勢圖已建立完成！", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillStockSourceData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("日期", "開盤價", "收盤價")
    ws.Range("A2:C2").Value = Array("2026/5/6",  150.5, 152.3)
    ws.Range("A3:C3").Value = Array("2026/5/7",  152.0, 149.8)
    ws.Range("A4:C4").Value = Array("2026/5/8",  149.5, 153.6)
    ws.Range("A5:C5").Value = Array("2026/5/9",  153.8, 155.2)
    ws.Range("A6:C6").Value = Array("2026/5/10", 154.9, 151.4)
    ws.Range("A2:A6").NumberFormat = "yyyy/m/d"
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateSheetStk(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetStk = ws
End Function

