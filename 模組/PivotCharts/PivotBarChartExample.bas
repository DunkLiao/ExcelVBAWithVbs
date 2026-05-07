Attribute VB_Name = "PivotBarChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotBarChartExample
'功能說明: 建立以樞紐分析表為資料來源的長條樞紐分析圖
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口
Sub TestPivotBarChart()
    Call CreatePivotBarChart
End Sub

' 建立樞紐分析圖（橫條圖）
Sub CreatePivotBarChart()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim cht As Chart

    Set wsData = GetOrCreateSheet(ThisWorkbook, "銷售資料")
    Call FillSalesData(wsData)

    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "樞紐分析圖")

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="銷售樞紐")

    With pt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With

    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum

    Set chartObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("E3").Left, _
        Top:=wsPivot.Range("E3").Top, _
        Width:=400, _
        Height:=280)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlBarClustered
    cht.HasTitle = True
    cht.ChartTitle.Text = "各地區銷售額樞紐分析圖"
    cht.ChartStyle = 5
    cht.SeriesCollection(1).HasDataLabels = True

    wsPivot.Activate
    MsgBox "樞紐分析圖已建立完成！", vbInformation, "完成"
End Sub

' 填入銷售資料
Private Sub FillSalesData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "產品"
    ws.Range("C1").Value = "銷售額"
    ws.Range("A2").Value = "北區"
    ws.Range("B2").Value = "產品A"
    ws.Range("C2").Value = 120000
    ws.Range("A3").Value = "南區"
    ws.Range("B3").Value = "產品A"
    ws.Range("C3").Value = 95000
    ws.Range("A4").Value = "東區"
    ws.Range("B4").Value = "產品B"
    ws.Range("C4").Value = 108000
    ws.Range("A5").Value = "西區"
    ws.Range("B5").Value = "產品B"
    ws.Range("C5").Value = 87000
    ws.Range("A6").Value = "北區"
    ws.Range("B6").Value = "產品B"
    ws.Range("C6").Value = 135000
    ws.Range("A7").Value = "南區"
    ws.Range("B7").Value = "產品A"
    ws.Range("C7").Value = 78000
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
