Attribute VB_Name = "HeatmapColumnFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: HeatmapColumnFormatting
'功能說明: 對每個欄位分別套用熱圖色階條件式格式，凸顯各欄位的高低分佈
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestHeatmapColumnFormatting()
    Dim ws As Worksheet
    Set ws = GetOrCreateHeatmapSheet(ThisWorkbook, "欄位熱圖格式範例")
    Call FillHeatmapData(ws)
    Call ApplyHeatmapColumnFormatting(ws)
    MsgBox "欄位熱圖格式套用完成！", vbInformation, "完成"
End Sub

Sub ApplyHeatmapColumnFormatting(ByVal ws As Worksheet)
    Dim lastRow      As Long
    Dim lastCol      As Long
    Dim colIdx       As Integer
    Dim rng          As Range
    Dim cfColorScale As ColorScale

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.UsedRange.Columns.Count

    If lastRow < 2 Then
        MsgBox "資料不足，無法套用格式。", vbExclamation, "警告"
        Exit Sub
    End If

    ws.UsedRange.FormatConditions.Delete

    For colIdx = 2 To lastCol
        Set rng = ws.Range(ws.Cells(2, colIdx), ws.Cells(lastRow, colIdx))
        rng.FormatConditions.AddColorScale ColorScaleType:=3
        Set cfColorScale = rng.FormatConditions(rng.FormatConditions.Count)

        With cfColorScale.ColorScaleCriteria(1)
            .Type = xlConditionValueLowestValue
            .FormatColor.Color = RGB(91, 155, 213)
        End With

        With cfColorScale.ColorScaleCriteria(2)
            .Type = xlConditionValuePercentile
            .Value = 50
            .FormatColor.Color = RGB(255, 255, 255)
        End With

        With cfColorScale.ColorScaleCriteria(3)
            .Type = xlConditionValueHighestValue
            .FormatColor.Color = RGB(255, 105, 97)
        End With
    Next colIdx
End Sub

Private Sub FillHeatmapData(ByVal ws As Worksheet)
    Dim dataArr As Variant
    Dim i       As Integer

    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("月份", "北區", "中區", "南區")

    dataArr = Array( _
        Array("一月", 85, 72, 91), _
        Array("二月", 78, 88, 65), _
        Array("三月", 92, 55, 83), _
        Array("四月", 61, 94, 79), _
        Array("五月", 88, 70, 96), _
        Array("六月", 74, 82, 68))

    For i = 0 To UBound(dataArr)
        ws.Cells(i + 2, 1).Value = dataArr(i)(0)
        ws.Cells(i + 2, 2).Value = dataArr(i)(1)
        ws.Cells(i + 2, 3).Value = dataArr(i)(2)
        ws.Cells(i + 2, 4).Value = dataArr(i)(3)
    Next i

    ws.Columns("A:D").AutoFit
End Sub

Private Function GetOrCreateHeatmapSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateHeatmapSheet = ws
End Function
