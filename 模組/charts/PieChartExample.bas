Attribute VB_Name = "PieChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: 圓餅圖範例
'功能描述: 在Excel中建立圓餅圖的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/3
'
'傳入參數:
'傳回資料:
'
'*************************************************************************************

' 簡化測試入口
Sub TestPieChart()
    Call CreatePieChart("圓餅圖範例")
End Sub

' 建立圓餅圖
' sheetName: 要建立圖表的工作表名稱
Sub CreatePieChart(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    ' 取得或建立工作表
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ' 清除工作表
    ws.Cells.Clear

    ' 建立範例資料
    Call FillPieChartData(ws)

    ' 設定資料範圍 (A1:B6)
    Set dataRange = ws.Range("A1:B6")

    ' 在工作表中新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=400, _
        Height:=320)

    Set cht = chartObj.Chart

    ' 設定圖表資料來源
    cht.SetSourceData Source:=dataRange

    ' 設定圖表類型為圓餅圖
    cht.ChartType = xlPie

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "產品市場佔有率分析"

    ' 顯示圖例
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight

    ' 顯示資料標籤（含百分比）
    With cht.SeriesCollection(1)
        .HasDataLabels = True
        With .DataLabels
            .ShowPercentage = True
            .ShowCategoryName = True
            .ShowValue = False
            .Separator = Chr(10)
        End With
    End With

    ' 設定圖表樣式
    cht.ChartStyle = 26

    MsgBox "圓餅圖已建立完成！", vbInformation, "完成"
End Sub

' 建立環圈圖（圓餅圖變體）
Sub CreateDonutChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range
    Dim sheetName As String

    sheetName = "環圈圖範例"

    ' 取得或建立工作表
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ' 清除工作表
    ws.Cells.Clear

    ' 建立範例資料
    Call FillDonutChartData(ws)

    ' 設定資料範圍
    Set dataRange = ws.Range("A1:B7")

    ' 新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=420, _
        Height:=340)

    Set cht = chartObj.Chart

    ' 設定圖表資料來源
    cht.SetSourceData Source:=dataRange

    ' 設定圖表類型為環圈圖
    cht.ChartType = xlDoughnut

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "部門預算分配比例"

    ' 顯示圖例
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    ' 顯示資料標籤
    With cht.SeriesCollection(1)
        .HasDataLabels = True
        With .DataLabels
            .ShowPercentage = True
            .ShowCategoryName = True
            .ShowValue = False
        End With
    End With

    ' 設定圖表樣式
    cht.ChartStyle = 26

    MsgBox "環圈圖已建立完成！", vbInformation, "完成"
End Sub

' 填入圓餅圖範例資料（產品市佔率）
Private Sub FillPieChartData(ByVal ws As Worksheet)
    ' 標題列
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "市占率"

    ' 資料列
    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 35

    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 25

    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 20

    ws.Range("A5").Value = "產品D"
    ws.Range("B5").Value = 12

    ws.Range("A6").Value = "其他"
    ws.Range("B6").Value = 8

    ' 自動調整欄寬
    ws.Columns("A:B").AutoFit
End Sub

' 填入環圈圖範例資料（部門預算）
Private Sub FillDonutChartData(ByVal ws As Worksheet)
    ' 標題列
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "預算（萬元）"

    ' 資料列
    ws.Range("A2").Value = "業務部"
    ws.Range("B2").Value = 500

    ws.Range("A3").Value = "研發部"
    ws.Range("B3").Value = 800

    ws.Range("A4").Value = "行銷部"
    ws.Range("B4").Value = 350

    ws.Range("A5").Value = "人資部"
    ws.Range("B5").Value = 200

    ws.Range("A6").Value = "財務部"
    ws.Range("B6").Value = 150

    ws.Range("A7").Value = "資訊部"
    ws.Range("B7").Value = 300

    ' 自動調整欄寬
    ws.Columns("A:B").AutoFit
End Sub