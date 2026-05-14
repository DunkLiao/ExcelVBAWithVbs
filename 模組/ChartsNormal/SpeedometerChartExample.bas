Attribute VB_Name = "SpeedometerChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: SpeedometerChartExample
'功能說明: 在 Excel 中以甜甜圈圖模擬速度計（儀表板）圖的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestSpeedometerChart()
    Call CreateSpeedometerChart("速度計圖範例")
End Sub

' 建立速度計圖（使用甜甜圈圖模擬）
' sheetName: 要建立圖的工作表名稱
Sub CreateSpeedometerChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim gaugeValue As Double
    Dim remaining As Double

    Set ws = GetOrCreateSpeedoWs(sheetName)
    ws.Cells.Clear

    ' 設定儀表數值（0~100，目前值為 72）
    gaugeValue = 72
    remaining = 100 - gaugeValue

    ' 填入甜甜圈資料（分為：已達成、未達成、下半部遮蔽區）
    ws.Range("A1").Value = "指標"
    ws.Range("B1").Value = "數值"
    ws.Range("A2").Value = "已達成"
    ws.Range("B2").Value = gaugeValue
    ws.Range("A3").Value = "未達成"
    ws.Range("B3").Value = remaining
    ws.Range("A4").Value = "遮蔽區"
    ws.Range("B4").Value = 100

    ws.Range("A6").Value = "目前數值"
    ws.Range("B6").Value = gaugeValue

    ws.Columns("A:B").AutoFit

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=400, _
        Height:=300)

    Set cht = chartObj.Chart
    cht.ChartType = xlDoughnut
    cht.SetSourceData Source:=ws.Range("A1:B4")

    cht.HasTitle = True
    cht.ChartTitle.Text = "業績達成率 " & gaugeValue & "%"
    cht.HasLegend = False

    ' 設定各段填色
    With cht.SeriesCollection(1)
        .Points(1).Format.Fill.ForeColor.RGB = RGB(0, 180, 0)
        .Points(2).Format.Fill.ForeColor.RGB = RGB(220, 220, 220)
        .Points(3).Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Points(3).Format.Line.Visible = msoFalse
        .DoughnutHoleSize = 60
    End With

    MsgBox "速度計圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立速度計圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateSpeedoWs(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSpeedoWs = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSpeedoWs Is Nothing Then
        Set GetOrCreateSpeedoWs = ThisWorkbook.Worksheets.Add
        GetOrCreateSpeedoWs.Name = sheetName
    End If
End Function
