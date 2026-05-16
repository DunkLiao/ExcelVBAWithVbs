Attribute VB_Name = "MixedAxisChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: MixedAxisChartExample
'功能說明: 建立雙Y軸混合圖表，左軸顯示銷售數量（群組直條圖），
'          右軸顯示達成率（折線圖），示範主次座標軸的設定方式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

Sub CreateMixedAxisChart()
    Dim ws         As Worksheet
    Dim chtObj     As ChartObject
    Dim oChart     As Chart
    Dim oSeries    As Series
    Dim sheetName  As String

    sheetName = "雙Y軸混合圖表"

    ' 取得或建立工作表
    Set ws = GetOrCreateSheet(sheetName)
    ws.Cells.Clear

    ' 填入範例資料
    Call FillMixedAxisData(ws)

    ' 移除舊圖表
    Dim obj As ChartObject
    For Each obj In ws.ChartObjects
        obj.Delete
    Next obj

    ' 插入圖表物件
    Set chtObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=480, Height:=300)

    Set oChart = chtObj.Chart

    ' 設定資料來源（A1:C7）
    oChart.SetSourceData Source:=ws.Range("A1:C7")
    oChart.ChartType = xlColumnClustered

    ' 第一數列：銷售數量，主座標軸（左Y軸），群組直條
    Set oSeries = oChart.SeriesCollection(1)
    oSeries.ChartType = xlColumnClustered
    oSeries.AxisGroup = xlPrimary

    ' 第二數列：達成率，次座標軸（右Y軸），折線
    Set oSeries = oChart.SeriesCollection(2)
    oSeries.ChartType = xlLine
    oSeries.AxisGroup = xlSecondary

    ' 主座標軸標題
    With oChart.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "銷售數量（件）"
    End With

    ' 次座標軸格式（百分比）
    With oChart.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = "達成率"
        .MinimumScale = 0
        .MaximumScale = 1.5
        .TickLabels.NumberFormat = "0%"
    End With

    ' 圖表標題
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "月銷售數量與達成率雙Y軸圖表"
    oChart.HasLegend = True

    MsgBox "雙Y軸混合圖表已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillMixedAxisData(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "月份"
    ws.Cells(1, 2).Value = "銷售數量"
    ws.Cells(1, 3).Value = "達成率"

    Dim r As Integer
    Dim months(1 To 6)  As String
    Dim qty(1 To 6)     As Long
    Dim rate(1 To 6)    As Double

    months(1) = "一月": qty(1) = 320: rate(1) = 0.8
    months(2) = "二月": qty(2) = 410: rate(2) = 1.02
    months(3) = "三月": qty(3) = 390: rate(3) = 0.97
    months(4) = "四月": qty(4) = 450: rate(4) = 1.12
    months(5) = "五月": qty(5) = 480: rate(5) = 1.2
    months(6) = "六月": qty(6) = 430: rate(6) = 1.07

    For r = 1 To 6
        ws.Cells(r + 1, 1).Value = months(r)
        ws.Cells(r + 1, 2).Value = qty(r)
        ws.Cells(r + 1, 3).Value = rate(r)
        ws.Cells(r + 1, 3).NumberFormat = "0%"
    Next r
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function
