Attribute VB_Name = "ClearChartFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearChartFormatting
'功能說明: 清除工作表中所有圖表物件的自訂格式設定，包含標題、背景色與框線，還原預設外觀
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口
Sub TestClearChartFormatting()
    Call ClearAllChartFormattingInSheet(ActiveSheet)
End Sub

' 清除指定工作表中所有圖表的格式
' ws: 目標工作表
Sub ClearAllChartFormattingInSheet(ByVal ws As Worksheet)
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim clearedCount As Integer

    If ws.ChartObjects.Count = 0 Then
        MsgBox "此工作表中沒有圖表物件。", vbInformation, "清除圖表格式"
        Exit Sub
    End If

    clearedCount = 0
    For Each chartObj In ws.ChartObjects
        Set cht = chartObj.Chart
        Call ResetSingleChartFormat(cht)
        clearedCount = clearedCount + 1
    Next chartObj

    MsgBox "已清除 " & clearedCount & " 個圖表的格式設定。", vbInformation, "清除完成"
End Sub

' 重置單一圖表格式為預設
' cht: 目標圖表物件
Sub ResetSingleChartFormat(ByVal cht As Chart)
    ' 清除圖表背景
    On Error Resume Next
    cht.PlotArea.Interior.ColorIndex = xlColorIndexNone
    cht.ChartArea.Interior.ColorIndex = xlColorIndexNone
    cht.ChartArea.Border.LineStyle = xlLineStyleNone

    ' 清除標題格式
    If cht.HasTitle Then
        With cht.ChartTitle
            .Font.ColorIndex = xlColorIndexAutomatic
            .Font.Bold = False
            .Font.Size = 12
        End With
    End If

    ' 清除各數列的自訂顏色
    Dim ser As Series
    Dim i As Integer
    For i = 1 To cht.SeriesCollection.Count
        Set ser = cht.SeriesCollection(i)
        ser.Interior.ColorIndex = xlColorIndexAutomatic
        ser.Border.ColorIndex = xlColorIndexAutomatic
    Next i

    ' 清除圖例格式
    If cht.HasLegend Then
        With cht.Legend
            .Interior.ColorIndex = xlColorIndexNone
            .Border.LineStyle = xlLineStyleNone
        End With
    End If
    On Error GoTo 0
End Sub