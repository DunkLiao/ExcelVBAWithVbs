Option Explicit
Attribute VB_Name = "MergeWithDataBarSummary"
'*************************************************************************************
'模組名稱: MergeWithDataBarSummary
'功能說明: 跨工作表合併資料並在合併結果中加入資料橫條作為摘要視覺化
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestMergeWithDataBar()
    Call MergeSheetsWithDataBarSummary
End Sub

' 合併所有工作表資料並加入資料橫條
Sub MergeSheetsWithDataBarSummary()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsSummary As Worksheet
    Dim wsName As String
    Dim dstRow As Long
    Dim srcLastRow As Long
    Dim i As Long
    Dim barRange As Range

    Set wb = ThisWorkbook
    wsName = "合併摘要含資料橫條"
    Set wsSummary = GetOrCreateWorksheet(wsName)
    wsSummary.Cells.Clear

    ' 設定摘要標題
    wsSummary.Range("A1").Value = "來源工作表"
    wsSummary.Range("B1").Value = "資料列數"
    dstRow = 2

    ' 逐一掃描所有工作表
    For Each ws In wb.Worksheets
        If ws.Name <> wsSummary.Name Then
            srcLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If srcLastRow >= 1 Then
                wsSummary.Cells(dstRow, 1).Value = ws.Name
                wsSummary.Cells(dstRow, 2).Value = srcLastRow
                dstRow = dstRow + 1
            End If
        End If
    Next ws

    If dstRow <= 2 Then
        MsgBox "沒有其他工作表可供合併。", vbExclamation, "提示"
        Exit Sub
    End If

    ' 在資料列數欄位加入資料橫條
    Set barRange = wsSummary.Range("B2:B" & (dstRow - 1))
    With barRange.FormatConditions.AddDatabar
        .BarColor.RGB = RGB(0, 112, 192)
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With

    wsSummary.Columns.AutoFit
    MsgBox "合併摘要完成，已加入資料橫條視覺化。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
