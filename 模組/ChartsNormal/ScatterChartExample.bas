Attribute VB_Name = "ScatterChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: 散佈圖範例
'功能描述: 在Excel中建立XY散佈圖的範例程式
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
Sub TestScatterChart()
    Call CreateScatterChart("散佈圖範例")
End Sub

' 建立XY散佈圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateScatterChart(ByVal sheetName As String)
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
    Call FillScatterData(ws)

    ' 設定資料範圍 (A1:B11)
    Set dataRange = ws.Range("A1:B11")

    ' 在工作表中新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=420, _
        Height:=320)

    Set cht = chartObj.Chart

    ' 設定圖表資料來源
    cht.SetSourceData Source:=dataRange

    ' 設定圖表類型為XY散佈圖（含資料標記）
    cht.ChartType = xlXYScatter

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "廣告費用與銷售額相關分析"

    ' 設定 X 軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "廣告費用（萬元）"
    End With

    ' 設定 Y 軸標題
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售額（萬元）"
    End With

    ' 設定圖表樣式
    cht.ChartStyle = 15

    ' 不顯示圖例（單系列不需要）
    cht.HasLegend = False

    MsgBox "散佈圖已建立完成！", vbInformation, "完成"
End Sub

' 建立含趨勢線的散佈圖
Sub CreateScatterWithTrendline()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range
    Dim ser As Series
    Dim tl As Trendline
    Dim sheetName As String

    sheetName = "散佈圖含趨勢線"

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
    Call FillScatterData(ws)

    ' 設定資料範圍
    Set dataRange = ws.Range("A1:B11")

    ' 新增圖表物件
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=480, _
        Height:=360)

    Set cht = chartObj.Chart

    ' 設定圖表資料來源
    cht.SetSourceData Source:=dataRange

    ' 設定圖表類型為XY散佈圖（含資料標記）
    cht.ChartType = xlXYScatter

    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "廣告費用與銷售額相關分析（含趨勢線）"

    ' 設定 X 軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "廣告費用（萬元）"
    End With

    ' 設定 Y 軸標題
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售額（萬元）"
    End With

    ' 取得第一個數列
    Set ser = cht.SeriesCollection(1)

    ' 新增線性趨勢線
    Set tl = ser.Trendlines.Add( _
        Type:=xlLinear, _
        Forward:=0, _
        Backward:=0, _
        DisplayEquation:=True, _
        DisplayRSquared:=True, _
        Name:="線性趨勢")

    ' 設定趨勢線格式
    With tl.Border
        .Color = RGB(255, 0, 0)
        .Weight = xlThin
    End With

    ' 設定圖表樣式
    cht.ChartStyle = 15
    cht.HasLegend = False

    MsgBox "含趨勢線的散佈圖已建立完成！", vbInformation, "完成"
End Sub

' 填入散佈圖範例資料（廣告費用 vs 銷售額）
Private Sub FillScatterData(ByVal ws As Worksheet)
    ' 標題列
    ws.Range("A1").Value = "廣告費用"
    ws.Range("B1").Value = "銷售額"

    ' 資料列（X值, Y值）
    ws.Range("A2").Value = 10
    ws.Range("B2").Value = 85

    ws.Range("A3").Value = 15
    ws.Range("B3").Value = 120

    ws.Range("A4").Value = 20
    ws.Range("B4").Value = 165

    ws.Range("A5").Value = 25
    ws.Range("B5").Value = 190

    ws.Range("A6").Value = 30
    ws.Range("B6").Value = 230

    ws.Range("A7").Value = 38
    ws.Range("B7").Value = 275

    ws.Range("A8").Value = 45
    ws.Range("B8").Value = 320

    ws.Range("A9").Value = 50
    ws.Range("B9").Value = 360

    ws.Range("A10").Value = 60
    ws.Range("B10").Value = 410

    ws.Range("A11").Value = 70
    ws.Range("B11").Value = 480

    ' 自動調整欄寬
    ws.Columns("A:B").AutoFit
End Sub