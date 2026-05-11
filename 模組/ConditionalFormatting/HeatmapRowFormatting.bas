Attribute VB_Name = "HeatmapRowFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: HeatmapRowFormatting
'功能說明: 依各列資料大小套用熱力圖色階，以顏色視覺化各儲存格數值強弱
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub ApplyHeatmapRowFormatting()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim cs As ColorScale

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < 2 Then
        MsgBox "資料範圍不足，請確認工作表有資料。", vbExclamation, "提示"
        Exit Sub
    End If

    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))

    dataRange.FormatConditions.Delete

    Set cs = dataRange.FormatConditions.AddColorScale(ColorScaleType:=3)

    cs.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 255, 255)

    cs.ColorScaleCriteria(2).Type = xlConditionValuePercentile
    cs.ColorScaleCriteria(2).Value = 50
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 200, 100)

    cs.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    cs.ColorScaleCriteria(3).FormatColor.Color = RGB(192, 0, 0)

    MsgBox "熱力圖色階條件式格式已套用完成！", vbInformation, "完成"
End Sub

Sub ClearHeatmapRowFormatting()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < 1 Then Exit Sub

    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).FormatConditions.Delete
    MsgBox "熱力圖格式已清除。", vbInformation, "完成"
End Sub
