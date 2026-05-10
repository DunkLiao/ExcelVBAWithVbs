'*************************************************************************************
'模組名稱: PivotFunnelChartExample
'功能說明: 根據樞紐分析表建立漏斗圖 (Funnel Chart) 範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************
Option Explicit

Sub CreatePivotFunnelChartExample()
    Dim ws          As Worksheet
    Dim wsChart     As Worksheet
    Dim pt          As PivotTable
    Dim chtObj      As ChartObject
    Dim cht         As Chart
    Dim pc          As PivotCache

    ' 建立資料工作表
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "FunnelSalesData"

    ws.Cells(1, 1).Value = "階段"
    ws.Cells(1, 2).Value = "金額"

    Dim data(1 To 5, 1 To 2) As Variant
    data(1, 1) = "開發": data(1, 2) = 8000000
    data(2, 1) = "提案": data(2, 2) = 5500000
    data(3, 1) = "議價": data(3, 2) = 3200000
    data(4, 1) = "簽約": data(4, 2) = 1800000
    data(5, 1) = "交貨": data(5, 2) = 900000

    Dim i As Integer
    For i = 1 To 5
        ws.Cells(i + 1, 1).Value = data(i, 1)
        ws.Cells(i + 1, 2).Value = data(i, 2)
    Next i

    ' 建立樞紐快取
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ws.Range("A1:B6"))

    ' 建立圖表工作表
    Set wsChart = ThisWorkbook.Sheets.Add(After:=ws)
    wsChart.Name = "PivotFunnelChart"

    ' 建立樞紐分析表
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsChart.Range("A3"), _
        TableName:="FunnelPT")

    With pt
        .PivotFields("階段").Orientation = xlRowField
        .PivotFields("階段").Position = 1
        .AddDataField .PivotFields("金額"), "加總 - 金額", xlSum
    End With

    ' 建立漏斗圖
    Set chtObj = wsChart.ChartObjects.Add(Left:=200, Top:=20, Width:=420, Height:=300)
    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlFunnel
    cht.HasTitle = True
    cht.ChartTitle.Text = "銷售漏斗樞紐圖"

    MsgBox "樞紐漏斗圖建立完成！", vbInformation, "完成"
End Sub
