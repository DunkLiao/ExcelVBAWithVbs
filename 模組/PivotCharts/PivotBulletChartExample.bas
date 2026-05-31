Attribute VB_Name = "PivotBulletChartExample"
Option Explicit

'*************************************************************************************
'模組名稱: PivotBulletChartExample
'功能說明: 依據工作表資料建立類子彈圖效果的組合圖
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub CreatePivotBulletChart()
    Dim ws As Worksheet
    Dim cht As ChartObject

    Set ws = ThisWorkbook.Worksheets(1)

    If ws.Cells(1, 1).Value = "" Then
        MsgBox "請確認 A1:C1 含有標題，A2 起為資料！", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set cht = ws.ChartObjects.Add( _
        Left:=250, Top:=20, Width:=420, Height:=260)

    With cht.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ws.Range("A1:C" & lastRow)
        .HasTitle = True
        .ChartTitle.Text = "子彈圖（Bullet Chart）"

        .SeriesCollection(2).ChartType = xlLine
        .SeriesCollection(2).MarkerStyle = xlMarkerStyleDiamond
        .SeriesCollection(2).MarkerSize = 8

        With .SeriesCollection(1)
            .Interior.Color = RGB(70, 130, 180)
        End With
    End With

    MsgBox "子彈圖已建立！", vbInformation
End Sub
