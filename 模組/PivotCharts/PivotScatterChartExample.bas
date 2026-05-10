Attribute VB_Name = "PivotScatterChartExample"
Option Explicit

' ============================================================
' 模組名稱：PivotScatterChartExample
' 功能說明：建立樞紐資料的散佈圖，呈現兩個數值欄位之間的關係
'           注意：Excel 不支援直接將樞紐分析表作為散佈圖來源，
'                 本範例先複製樞紐資料為純數值，再建立散佈圖
' ============================================================

Sub CreatePivotScatterChartExample()
    Dim wb          As Workbook
    Dim wsSrc       As Worksheet
    Dim wsPvt       As Worksheet
    Dim wsChart     As Worksheet
    Dim pvtCache    As PivotCache
    Dim pvt         As PivotTable
    Dim chtObj      As ChartObject
    Dim cht         As Chart
    Dim rngData     As Range
    Dim lastRow     As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Set wb = ThisWorkbook
    
    ' 建立範例資料工作表
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("散佈圖資料").Delete
    wb.Sheets("散佈圖樞紐").Delete
    wb.Sheets("散佈圖").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    ' --- 建立原始資料 ---
    Set wsSrc = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsSrc.Name = "散佈圖資料"
    
    wsSrc.Range("A1:C1").Value = Array("業務員", "拜訪次數", "成交金額(萬)")
    wsSrc.Rows(1).Font.Bold = True
    
    Dim salesData(9, 3) As Variant
    Dim salesNames(9) As String
    Dim visits(9) As Integer
    Dim amounts(9) As Double
    
    salesNames(0) = "王大明" : visits(0) = 20 : amounts(0) = 85.5
    salesNames(1) = "李小芳" : visits(1) = 35 : amounts(1) = 142.0
    salesNames(2) = "張志遠" : visits(2) = 15 : amounts(2) = 62.3
    salesNames(3) = "陳美玲" : visits(3) = 42 : amounts(3) = 178.9
    salesNames(4) = "林建國" : visits(4) = 28 : amounts(4) = 110.2
    salesNames(5) = "黃淑芬" : visits(5) = 10 : amounts(5) = 41.0
    salesNames(6) = "吳宗翰" : visits(6) = 50 : amounts(6) = 210.5
    salesNames(7) = "蔡雅婷" : visits(7) = 22 : amounts(7) = 95.8
    salesNames(8) = "許建宏" : visits(8) = 38 : amounts(8) = 155.0
    salesNames(9) = "葉欣儀" : visits(9) = 30 : amounts(9) = 120.7
    
    Dim i As Integer
    For i = 0 To 9
        wsSrc.Cells(i + 2, 1).Value = salesNames(i)
        wsSrc.Cells(i + 2, 2).Value = visits(i)
        wsSrc.Cells(i + 2, 3).Value = amounts(i)
    Next i
    wsSrc.Columns.AutoFit
    
    ' --- 建立樞紐分析表 ---
    Set wsPvt = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsPvt.Name = "散佈圖樞紐"
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    Set pvtCache = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsSrc.Range("A1:C" & lastRow))
    
    Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=wsPvt.Range("A1"), _
        TableName:="PvtScatter")
    
    With pvt.PivotFields("業務員")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pvt.PivotFields("拜訪次數")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "拜訪次數合計"
    End With
    With pvt.PivotFields("成交金額(萬)")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "成交金額合計"
    End With
    
    ' --- 將樞紐結果複製為純數值，供散佈圖使用 ---
    Set wsChart = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsChart.Name = "散佈圖"
    
    wsChart.Range("A1").Value = "業務員"
    wsChart.Range("B1").Value = "拜訪次數"
    wsChart.Range("C1").Value = "成交金額(萬)"
    wsChart.Rows(1).Font.Bold = True
    
    Dim pvtRow As Integer
    Dim destRow As Integer
    destRow = 2
    pvtRow = pvt.DataBodyRange.Row
    
    Dim pvtLastRow As Long
    pvtLastRow = pvt.DataBodyRange.Row + pvt.DataBodyRange.Rows.Count - 1
    
    Dim r As Long
    For r = pvtRow To pvtLastRow
        Dim itmLabel As String
        itmLabel = CStr(wsPvt.Cells(r, 1).Value)
        If itmLabel <> "" And itmLabel <> "總計" Then
            wsChart.Cells(destRow, 1).Value = itmLabel
            wsChart.Cells(destRow, 2).Value = wsPvt.Cells(r, 2).Value
            wsChart.Cells(destRow, 3).Value = wsPvt.Cells(r, 3).Value
            destRow = destRow + 1
        End If
    Next r
    wsChart.Columns.AutoFit
    
    ' --- 建立散佈圖 ---
    Dim chartLastRow As Long
    chartLastRow = wsChart.Cells(wsChart.Rows.Count, 1).End(xlUp).Row
    Set rngData = wsChart.Range("B1:C" & chartLastRow)
    
    Set chtObj = wsChart.ChartObjects.Add(Left:=10, Top:=160, Width:=500, Height:=320)
    Set cht = chtObj.Chart
    
    cht.ChartType = xlXYScatterSmooth
    cht.SetSourceData Source:=rngData
    
    cht.HasTitle = True
    cht.ChartTitle.Text = "業務員拜訪次數 vs 成交金額散佈圖"
    
    cht.Axes(xlCategory).HasTitle = True
    cht.Axes(xlCategory).AxisTitle.Text = "拜訪次數"
    cht.Axes(xlValue).HasTitle = True
    cht.Axes(xlValue).AxisTitle.Text = "成交金額（萬元）"
    
    cht.HasLegend = False
    
    ' 加入趨勢線
    cht.SeriesCollection(1).Trendlines.Add( _
        Type:=xlLinear, _
        DisplayEquation:=True, _
        DisplayRSquared:=True)
    
    Application.ScreenUpdating = True
    MsgBox "樞紐散佈圖建立完成！請查看「散佈圖」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub