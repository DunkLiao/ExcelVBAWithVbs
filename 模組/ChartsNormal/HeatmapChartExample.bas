Option Explicit
Attribute VB_Name = "HeatmapChartExample"
'*************************************************************************************

'模組名稱: HeatmapChartExample

'功能說明: 使用條件式格式模擬熱力圖（Heatmap），以色階呈現數值分佈

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub CreateHeatmapChartExample()

    Dim ws As Worksheet

    Dim rngData As Range

    Dim i As Integer

    Dim j As Integer

    Dim rowCount As Integer

    Dim colCount As Integer



    rowCount = 8

    colCount = 8



    ' 建立新工作表

    Set ws = ThisWorkbook.Worksheets.Add

    ws.Name = "HeatmapExample"



    ' 填入欄位標題

    For j = 1 To colCount

        ws.Cells(1, j + 1).Value = "欄" & j

    Next j

    For i = 1 To rowCount

        ws.Cells(i + 1, 1).Value = "列" & i

    Next i



    ' 填入隨機數值（1~100）

    Randomize

    For i = 1 To rowCount

        For j = 1 To colCount

            ws.Cells(i + 1, j + 1).Value = Int(Rnd() * 100) + 1

        Next j

    Next i



    ' 設定資料範圍

    Set rngData = ws.Range(ws.Cells(2, 2), ws.Cells(rowCount + 1, colCount + 1))



    ' 套用色階條件格式（綠-黃-紅）

    rngData.FormatConditions.Delete

    With rngData.FormatConditions.AddColorScale(ColorScaleType:=3)

        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue

        .ColorScaleCriteria(1).FormatColor.Color = RGB(87, 187, 138)

        .ColorScaleCriteria(2).Type = xlConditionValuePercentile

        .ColorScaleCriteria(2).Value = 50

        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)

        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue

        .ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)

    End With



    ' 自動調整欄寬與列高

    rngData.EntireColumn.AutoFit

    rngData.EntireRow.RowHeight = 22



    ' 設定儲存格對齊方式

    With rngData

        .HorizontalAlignment = xlCenter

        .VerticalAlignment = xlCenter

    End With



    ' 加入標題

    ws.Range("A1").Value = "熱力圖範例"

    With ws.Range("A1")

        .Font.Bold = True

        .Font.Size = 13

    End With



    MsgBox "熱力圖已建立完成，工作表：" & ws.Name, vbInformation, "完成"

End Sub

