Attribute VB_Name = "BarChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: 長條圖範例
'功能描述: 在Excel中建立長條圖的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/3
'
'傳入參數:
'傳回數值:
'
'*************************************************************************************

' 測試用入口
Sub TestBarChart()
    Call CreateBarChart("長條圖範例")
End Sub

' 建立長條圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateBarChart(ByVal sheetName As String)
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
    Call FillSampleData(ws)

    ' 設定資料範圍 (A1:B6)
    Set dataRange = ws.Range("A1:B6")

    ' 在工作表中新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=400, _
        Height:=300)

    Set cht = chartObj.Chart

    ' 設定圖表資料來源
    cht.SetSourceData Source:=dataRange

    ' 設定圖表類型為群組直條圖
    cht.ChartType = xlColumnClustered

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "月份銷售量統計"

    ' 設定 X 軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "月份"
    End With

    ' 設定 Y 軸標題
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售量"
    End With

    ' 設定圖表樣式
    cht.ChartStyle = 2

    ' 顯示資料標籤
    cht.SeriesCollection(1).HasDataLabels = True

    MsgBox "長條圖已建立完成！", vbInformation, "完成"
End Sub

' 填入範例資料
Private Sub FillSampleData(ByVal ws As Worksheet)
    ' 標題列
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "銷售量"

    ' 資料列
    ws.Range("A2").Value = "一月"
    ws.Range("B2").Value = 1200

    ws.Range("A3").Value = "二月"
    ws.Range("B3").Value = 850

    ws.Range("A4").Value = "三月"
    ws.Range("B4").Value = 1560

    ws.Range("A5").Value = "四月"
    ws.Range("B5").Value = 970

    ws.Range("A6").Value = "五月"
    ws.Range("B6").Value = 1380

    ' 自動調整欄寬
    ws.Columns("A:B").AutoFit
End Sub

' 建立多系列長條圖範例
Sub CreateMultiSeriesBarChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range
    Dim sheetName As String

    sheetName = "多系列長條圖"

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

    ' 建立多系列範例資料
    Call FillMultiSeriesData(ws)

    ' 設定資料範圍
    Set dataRange = ws.Range("A1:C6")

    ' 新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=480, _
        Height:=320)

    Set cht = chartObj.Chart

    ' 設定圖表資料來源
    cht.SetSourceData Source:=dataRange

    ' 設定圖表類型為群組直條圖
    cht.ChartType = xlColumnClustered

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "各季度產品銷售比較"

    ' 設定 X 軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "季度"
    End With

    ' 設定 Y 軸標題
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售額（萬元）"
    End With

    ' 設定圖表樣式
    cht.ChartStyle = 10

    ' 顯示圖例
    cht.HasLegend = True

    MsgBox "多系列長條圖已建立完成！", vbInformation, "完成"
End Sub

' 填入多系列範例資料
Private Sub FillMultiSeriesData(ByVal ws As Worksheet)
    ' 標題列
    ws.Range("A1").Value = "季度"
    ws.Range("B1").Value = "產品A"
    ws.Range("C1").Value = "產品B"

    ' 資料列
    ws.Range("A2").Value = "第一季"
    ws.Range("B2").Value = 320
    ws.Range("C2").Value = 280

    ws.Range("A3").Value = "第二季"
    ws.Range("B3").Value = 450
    ws.Range("C3").Value = 390

    ws.Range("A4").Value = "第三季"
    ws.Range("B4").Value = 380
    ws.Range("C4").Value = 420

    ws.Range("A5").Value = "第四季"
    ws.Range("B5").Value = 510
    ws.Range("C5").Value = 480

    ws.Range("A6").Value = "全年合計"
    ws.Range("B6").Value = 1660
    ws.Range("C6").Value = 1570

    ' 自動調整欄寬
    ws.Columns("A:C").AutoFit
End Sub