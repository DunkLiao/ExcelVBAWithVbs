Option Explicit
Attribute VB_Name = "PivotHeatmapChartExample"
'*************************************************************************************
'模組名稱: 樞紐熱圖圖表範例
'功能說明: 以樞紐分析表為資料來源，建立色階熱圖（ColorScale 條件格式）
'          視覺化各區域各月份的銷售表現
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub TestPivotHeatmapChart()
    Call CreatePivotHeatmapChart("熱圖範例")
End Sub

Sub CreatePivotHeatmapChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache
    Dim dataRange As Range
    Dim pivotRange As Range

    Set wsData = GetOrCreateWsHeatmap(sheetName & "_資料")
    wsData.Cells.Clear
    Call FillHeatmapData(wsData)

    Set wsPivot = GetOrCreateWsHeatmap(sheetName & "_樞紐熱圖")
    wsPivot.Cells.Clear

    Set dataRange = wsData.Range("A1:C25")
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="HeatmapPivot")

    With pt
        .PivotFields("區域").Orientation = xlRowField
        .PivotFields("區域").Position = 1
        .PivotFields("月份").Orientation = xlColumnField
        .PivotFields("月份").Position = 1
        .AddDataField .PivotFields("銷售額"), "加總-銷售額", xlSum
    End With

    ' 找到數值範圍並加入色階條件格式
    Set pivotRange = pt.DataBodyRange
    If Not pivotRange Is Nothing Then
        pivotRange.FormatConditions.Delete

        Dim cs As ColorScale
        Set cs = pivotRange.FormatConditions.AddColorScale(ColorScaleType:=3)

        cs.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        cs.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 100, 100) ' 低值：紅

        cs.ColorScaleCriteria(2).Type = xlConditionValuePercentile
        cs.ColorScaleCriteria(2).Value = 50
        cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 153) ' 中值：黃

        cs.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        cs.ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123) ' 高值：綠
    End If

    wsPivot.Columns.AutoFit
    MsgBox "樞紐熱圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立樞紐熱圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetOrCreateWsHeatmap(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWsHeatmap = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWsHeatmap Is Nothing Then
        Set GetOrCreateWsHeatmap = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateWsHeatmap.Name = sheetName
    End If
End Function

Private Sub FillHeatmapData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "區域"
    ws.Range("B1").Value = "月份"
    ws.Range("C1").Value = "銷售額"

    Dim areas As Variant
    Dim months As Variant
    Dim vals As Variant

    areas = Array("北區", "北區", "北區", "北區", "北區", "北區", _
                  "中區", "中區", "中區", "中區", "中區", "中區", _
                  "南區", "南區", "南區", "南區", "南區", "南區", _
                  "東區", "東區", "東區", "東區", "東區", "東區")
    months = Array("一月", "二月", "三月", "四月", "五月", "六月", _
                   "一月", "二月", "三月", "四月", "五月", "六月", _
                   "一月", "二月", "三月", "四月", "五月", "六月", _
                   "一月", "二月", "三月", "四月", "五月", "六月")
    vals = Array(120, 95, 135, 160, 145, 180, _
                 88, 102, 115, 130, 122, 140, _
                 75, 90, 105, 98, 118, 135, _
                 60, 72, 85, 95, 110, 125)

    Dim i As Integer
    For i = 0 To 23
        ws.Cells(i + 2, 1).Value = areas(i)
        ws.Cells(i + 2, 2).Value = months(i)
        ws.Cells(i + 2, 3).Value = vals(i)
    Next i

    ws.Columns("A:C").AutoFit
End Sub
