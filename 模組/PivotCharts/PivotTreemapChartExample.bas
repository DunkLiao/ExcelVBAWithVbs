Attribute VB_Name = "PivotTreemapChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotTreemapChartExample
'功能說明: 建立樞紐分析表後，以矩形樹狀圖 (Treemap) 呈現各分類銷售佔比的範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestPivotTreemapChart()
    Call CreatePivotTreemapChart
End Sub

' 建立樞紐矩形樹狀圖
Sub CreatePivotTreemapChart()
    On Error GoTo ErrorHandler

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range
    Dim lastRow As Long
    Dim r As Integer

    ' 準備資料
    Set wsData = GetOrCreateWorksheetPTC("矩形樹狀資料")
    wsData.Cells.Clear
    Call FillTreemapData(wsData)

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set dataRange = wsData.Range("A1").Resize(lastRow, 3)

    ' 準備輸出工作表
    Set wsPivot = GetOrCreateWorksheetPTC("矩形樹狀樞紐圖")
    wsPivot.Cells.Clear

    ' 建立樞紐快取與分析表
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="矩形樹狀樞紐")

    With pt.PivotFields("類別")
        .Orientation = xlRowField
        .Position = 1
    End With

    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum

    ' 建立矩形樹狀圖
    Set chartObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("E3").Left, _
        Top:=wsPivot.Range("E3").Top, _
        Width:=500, _
        Height:=360)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlTreemap

    cht.HasTitle = True
    cht.ChartTitle.Text = "各產品類別銷售佔比矩形樹狀圖"

    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    wsPivot.Columns("A:D").AutoFit
    wsPivot.Activate

    MsgBox "樞紐矩形樹狀圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立矩形樹狀圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetOrCreateWorksheetPTC(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheetPTC = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateWorksheetPTC Is Nothing Then
        Set GetOrCreateWorksheetPTC = ThisWorkbook.Worksheets.Add
        GetOrCreateWorksheetPTC.Name = sheetName
    End If
End Function

Private Sub FillTreemapData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "類別"
    ws.Range("B1").Value = "產品"
    ws.Range("C1").Value = "銷售額"

    Dim r As Integer
    Dim data(1 To 10, 1 To 3) As Variant
    data(1, 1) = "電子" : data(1, 2) = "手機" : data(1, 3) = 550000
    data(2, 1) = "電子" : data(2, 2) = "平板" : data(2, 3) = 320000
    data(3, 1) = "電子" : data(3, 2) = "筆電" : data(3, 3) = 480000
    data(4, 1) = "服飾" : data(4, 2) = "上衣" : data(4, 3) = 120000
    data(5, 1) = "服飾" : data(5, 2) = "褲子" : data(5, 3) = 95000
    data(6, 1) = "服飾" : data(6, 2) = "鞋子" : data(6, 3) = 140000
    data(7, 1) = "食品" : data(7, 2) = "飲料" : data(7, 3) = 88000
    data(8, 1) = "食品" : data(8, 2) = "零食" : data(8, 3) = 72000
    data(9, 1) = "家居" : data(9, 2) = "家具" : data(9, 3) = 260000
    data(10, 1) = "家居" : data(10, 2) = "寢具" : data(10, 3) = 180000

    For r = 1 To 10
        ws.Cells(r + 1, 1).Value = data(r, 1)
        ws.Cells(r + 1, 2).Value = data(r, 2)
        ws.Cells(r + 1, 3).Value = data(r, 3)
    Next r

    ws.Columns("A:C").AutoFit
End Sub
