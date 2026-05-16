Attribute VB_Name = "MergeWithConditionalFormat"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithConditionalFormat
'功能說明: 合併活頁簿中所有工作表（排除目標工作表）的資料至彙總表，
'          並在合併完成後自動套用條件式格式，依數值高低標示顏色
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

Sub MergeSheetsAndApplyConditionalFormat()
    Dim wbThis      As Workbook
    Dim wsDest      As Worksheet
    Dim wsLoop      As Worksheet
    Dim destName    As String
    Dim destRow     As Long
    Dim srcLastRow  As Long
    Dim srcLastCol  As Long
    Dim hasHeader   As Boolean
    Dim numColIdx   As Long

    destName = "彙總資料"
    Set wbThis = ThisWorkbook

    ' 取得或建立彙總工作表
    Set wsDest = GetOrCreateDestSheet(wbThis, destName)
    wsDest.Cells.Clear
    destRow = 1
    hasHeader = False

    ' 逐一合併其他工作表
    For Each wsLoop In wbThis.Worksheets
        If wsLoop.Name <> destName Then
            srcLastRow = wsLoop.Cells(wsLoop.Rows.Count, 1).End(xlUp).Row
            srcLastCol = wsLoop.Cells(1, wsLoop.Columns.Count).End(xlToLeft).Column

            If srcLastRow >= 2 And srcLastCol >= 1 Then
                If Not hasHeader Then
                    wsLoop.Range(wsLoop.Cells(1, 1), wsLoop.Cells(srcLastRow, srcLastCol)) _
                          .Copy wsDest.Cells(destRow, 1)
                    destRow = destRow + srcLastRow
                    hasHeader = True
                Else
                    wsLoop.Range(wsLoop.Cells(2, 1), wsLoop.Cells(srcLastRow, srcLastCol)) _
                          .Copy wsDest.Cells(destRow, 1)
                    destRow = destRow + srcLastRow - 1
                End If
            End If
        End If
    Next wsLoop

    If destRow = 1 Then
        MsgBox "沒有找到可合併的工作表資料。", vbExclamation
        Exit Sub
    End If

    wsDest.Columns.AutoFit

    ' 找出第一個數值欄（第2欄開始搜尋）
    numColIdx = 0
    Dim c As Long
    Dim totalCols As Long
    totalCols = wsDest.Cells(1, wsDest.Columns.Count).End(xlToLeft).Column
    For c = 2 To totalCols
        If IsNumeric(wsDest.Cells(2, c).Value) And wsDest.Cells(2, c).Value <> "" Then
            numColIdx = c
            Exit For
        End If
    Next c

    ' 套用條件式格式（三色色階）
    If numColIdx > 0 Then
        Dim totalRows As Long
        totalRows = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row
        If totalRows > 1 Then
            Call ApplyThreeColorScale(wsDest, numColIdx, 2, totalRows)
        End If
    End If

    Dim dataCount As Long
    dataCount = destRow - 2
    If dataCount < 0 Then dataCount = 0
    MsgBox "合併完成！共 " & dataCount & " 筆資料，條件式格式已套用。", vbInformation, "完成"
End Sub

' 在指定欄套用三色色階條件式格式
Private Sub ApplyThreeColorScale(ByVal ws As Worksheet, ByVal colIdx As Long, _
                                   ByVal startRow As Long, ByVal endRow As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, colIdx), ws.Cells(endRow, colIdx))

    rng.FormatConditions.Delete

    Dim cf As ColorScale
    Set cf = rng.FormatConditions.AddColorScale(ColorScaleType:=3)

    ' 最小值：紅色
    cf.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    cf.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 99, 71)

    ' 中間值：黃色
    cf.ColorScaleCriteria(2).Type = xlConditionValuePercentile
    cf.ColorScaleCriteria(2).Value = 50
    cf.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0)

    ' 最大值：綠色
    cf.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    cf.ColorScaleCriteria(3).FormatColor.Color = RGB(0, 200, 80)
End Sub

Private Function GetOrCreateDestSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateDestSheet = ws
End Function
