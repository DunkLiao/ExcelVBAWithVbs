Attribute VB_Name = "ThreeDColumnChartExample"
Option Explicit

'*************************************************************************************
'模組名稱: ThreeDColumnChartExample
'功能說明: 建立三維群組直條圖範例
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub CreateThreeDColumnChart()
    '建立三維群組直條圖，使用工作表中的資料作為來源
    Dim ws As Worksheet
    Dim cht As ChartObject
    Dim rngData As Range

    Set ws = ThisWorkbook.Worksheets(1)
    Set rngData = ws.Range("A1:D5")

    Set cht = ws.ChartObjects.Add( _
        Left:=200, Top:=20, Width:=400, Height:=250)

    With cht.Chart
        .SetSourceData Source:=rngData, PlotBy:=xlColumns
        .ChartType = xl3DColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "三維群組直條圖"
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "數值"
        End With
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "類別"
        End With
    End With

    MsgBox "三維群組直條圖建立完成！", vbInformation
End Sub
