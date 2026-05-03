Attribute VB_Name = "LineChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: 折線圖範例
'功能描述: 在Excel中建立折線圖的範例程式
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
Sub TestLineChart()
    Call CreateLineChart("折線圖範例")
End Sub

' 建立單系列折線圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateLineChart(ByVal sheetName As String)
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
    Call FillLineChartData(ws)

    ' 設定資料範圍 (A1:B13)
    Set dataRange = ws.Range("A1:B13")

    ' 在工作表中新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=420, _
        Height:=300)

    Set cht = chartObj.Chart

    ' 設定圖表資料來源
    cht.SetSourceData Source:=dataRange

    ' 設定圖表類型為含資料標記的折線圖
    cht.ChartType = xlLineMarkers

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "全年每月氣溫趨勢"

    ' 設定 X 軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "月份"
    End With

    ' 設定 Y 軸標題
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "平均氣溫 (°C)"
        .MinimumScale = 0
        .MaximumScale = 40
    End With

    ' 設定圖表樣式
    cht.ChartStyle = 4

    ' 顯示資料標籤
    cht.SeriesCollection(1).HasDataLabels = True

    ' 設定折線平滑化
    cht.SeriesCollection(1).Smooth = True

    MsgBox "折線圖已建立完成！", vbInformation, "完成"
End Sub

' 建立雙系列折線圖（年度比較）
Sub CreateCompareLineChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range
    Dim sheetName As String

    sheetName = "折線圖年度比較"

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

    ' 建立雙系列範例資料
    Call FillCompareLineData(ws)

    ' 設定資料範圍
    Set dataRange = ws.Range("A1:C7")

    ' 新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=480, _
        Height:=320)

    Set cht = chartObj.Chart

    ' 設定圖表資料來源
    cht.SetSourceData Source:=dataRange

    ' 設定圖表類型
    cht.ChartType = xlLineMarkers

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "上下半年業績趨勢比較"

    ' 設定 X 軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "季度"
    End With

    ' 設定 Y 軸標題
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "業績（萬元）"
    End With

    ' 設定圖表樣式
    cht.ChartStyle = 12

    ' 顯示圖例
    cht.HasLegend = True

    MsgBox "雙系列折線圖已建立完成！", vbInformation, "完成"
End Sub

' 填入折線圖範例資料（月份氣溫）
Private Sub FillLineChartData(ByVal ws As Worksheet)
    ' 標題列
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "平均氣溫"

    ' 資料列
    ws.Range("A2").Value = "1月"
    ws.Range("B2").Value = 16

    ws.Range("A3").Value = "2月"
    ws.Range("B3").Value = 17

    ws.Range("A4").Value = "3月"
    ws.Range("B4").Value = 20

    ws.Range("A5").Value = "4月"
    ws.Range("B5").Value = 24

    ws.Range("A6").Value = "5月"
    ws.Range("B6").Value = 28

    ws.Range("A7").Value = "6月"
    ws.Range("B7").Value = 32

    ws.Range("A8").Value = "7月"
    ws.Range("B8").Value = 34

    ws.Range("A9").Value = "8月"
    ws.Range("B9").Value = 34

    ws.Range("A10").Value = "9月"
    ws.Range("B10").Value = 30

    ws.Range("A11").Value = "10月"
    ws.Range("B11").Value = 26

    ws.Range("A12").Value = "11月"
    ws.Range("B12").Value = 22

    ws.Range("A13").Value = "12月"
    ws.Range("B13").Value = 18

    ' 自動調整欄寬
    ws.Columns("A:B").AutoFit
End Sub

' 填入雙系列折線圖範例資料
Private Sub FillCompareLineData(ByVal ws As Worksheet)
    ' 標題列
    ws.Range("A1").Value = "季度"
    ws.Range("B1").Value = "上半年"
    ws.Range("C1").Value = "下半年"

    ' 資料列
    ws.Range("A2").Value = "第一季"
    ws.Range("B2").Value = 280
    ws.Range("C2").Value = 310

    ws.Range("A3").Value = "第二季"
    ws.Range("B3").Value = 350
    ws.Range("C3").Value = 420

    ws.Range("A4").Value = "第三季"
    ws.Range("B4").Value = 410
    ws.Range("C4").Value = 390

    ws.Range("A5").Value = "第四季"
    ws.Range("B5").Value = 380
    ws.Range("C5").Value = 450

    ws.Range("A6").Value = "第五季"
    ws.Range("B6").Value = 460
    ws.Range("C6").Value = 500

    ws.Range("A7").Value = "第六季"
    ws.Range("B7").Value = 520
    ws.Range("C7").Value = 480

    ' 自動調整欄寬
    ws.Columns("A:C").AutoFit
End Sub