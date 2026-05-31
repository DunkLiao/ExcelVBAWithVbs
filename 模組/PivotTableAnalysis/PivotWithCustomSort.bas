Attribute VB_Name = "PivotWithCustomSort"
Option Explicit

'*************************************************************************************
'模組名稱: PivotWithCustomSort
'功能說明: 建立樞紐分析表並套用自訂欄位排序順序
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub CreatePivotWithCustomSort()
    Dim wsSrc As Worksheet
    Dim wsPvt As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim srcRange As Range
    Dim lastRow As Long
    Dim lastCol As Long

    Set wsSrc = ThisWorkbook.Worksheets(1)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    Set srcRange = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol))

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("樞紐分析").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsPvt = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsPvt.Name = "樞紐分析"

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("A3"), _
        TableName:="CustomSortPivot")

    With pt
        Set pf = .PivotFields(wsSrc.Cells(1, 1).Value)
        pf.Orientation = xlRowField
        pf.Position = 1
        pf.AutoSort xlManual, pf.Name

        With .PivotFields(wsSrc.Cells(1, 2).Value)
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0"
        End With
    End With

    wsPvt.Columns("A:D").AutoFit
    MsgBox "樞紐分析表已建立，列欄位已設為手動排序！", vbInformation
End Sub
