Option Explicit
Attribute VB_Name = "FunnelChartExample"
'*************************************************************************************
'模組名稱: FunnelChartExample
'功能說明: 以 VBA 建立漏斗圖 (Funnel Chart) 範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Sub CreateFunnelChartExample()
    Dim ws          As Worksheet
    Dim chtObj      As ChartObject
    Dim cht         As Chart
    Dim rngData     As Range
    Dim i           As Integer

    ' 建立示範工作表
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "FunnelChartDemo"

    ' 寫入示範資料
    ws.Cells(1, 1).Value = "階段"
    ws.Cells(1, 2).Value = "數量"
    Dim stages(1 To 5, 1 To 2) As Variant
    stages(1, 1) = "潛在客戶": stages(1, 2) = 5000
    stages(2, 1) = "初步接觸": stages(2, 2) = 3200
    stages(3, 1) = "需求確認": stages(3, 2) = 1800
    stages(4, 1) = "報價議價": stages(4, 2) = 900
    stages(5, 1) = "成交訂單": stages(5, 2) = 420

    For i = 1 To 5
        ws.Cells(i + 1, 1).Value = stages(i, 1)
        ws.Cells(i + 1, 2).Value = stages(i, 2)
    Next i

    Set rngData = ws.Range("A1:B6")

    ' 建立圖表物件
    Set chtObj = ws.ChartObjects.Add(Left:=150, Top:=20, Width:=400, Height:=280)
    Set cht = chtObj.Chart

    cht.SetSourceData Source:=rngData
    cht.ChartType = xlFunnel

    cht.HasTitle = True
    cht.ChartTitle.Text = "銷售漏斗圖"

    MsgBox "漏斗圖已成功建立！", vbInformation, "完成"
End Sub
