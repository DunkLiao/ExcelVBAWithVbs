Attribute VB_Name = "MergeWithHyperlinkIndex"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithHyperlinkIndex
'功能說明: 合併所有非摘要工作表資料並建立含超連結的索引工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunMergeWithHyperlinkIndex()
    On Error GoTo ErrorHandler

    Dim wsSummary As Worksheet
    Dim wsIndex As Worksheet
    Dim ws As Worksheet
    Dim destRow As Long
    Dim indexRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim hasHeader As Boolean

    Set wsSummary = GetOrCreateMergeSheet("彙整")
    Set wsIndex = GetOrCreateMergeSheet("索引")

    wsSummary.Cells.Clear
    wsIndex.Cells.Clear
    wsIndex.Range("A1:C1").Value = Array("工作表", "資料列數", "說明")

    destRow = 1
    indexRow = 2
    hasHeader = False

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSummary.Name And ws.Name <> wsIndex.Name Then
            lastRow = GetMergeLastRow(ws)
            lastCol = GetMergeLastCol(ws)

            wsIndex.Hyperlinks.Add _
                Anchor:=wsIndex.Cells(indexRow, 1), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            wsIndex.Cells(indexRow, 2).Value = IIf(lastRow > 1, lastRow - 1, 0)
            wsIndex.Cells(indexRow, 3).Value = "按一下可跳至來源工作表"
            indexRow = indexRow + 1

            If lastRow >= 1 And lastCol >= 1 Then
                If Not hasHeader Then
                    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy _
                        Destination:=wsSummary.Cells(destRow, 1)
                    destRow = destRow + 1
                    hasHeader = True
                End If

                If lastRow > 1 Then
                    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy _
                        Destination:=wsSummary.Cells(destRow, 1)
                    destRow = destRow + lastRow - 1
                End If
            End If
        End If
    Next ws

    Application.ScreenUpdating = True
    wsSummary.Columns.AutoFit
    wsIndex.Columns.AutoFit

    MsgBox "已完成工作表合併並建立超連結索引。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "建立彙整與索引時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetMergeLastRow(ByVal ws As Worksheet) As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetMergeLastRow = 0
    Else
        GetMergeLastRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
End Function

Private Function GetMergeLastCol(ByVal ws As Worksheet) As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetMergeLastCol = 0
    Else
        GetMergeLastCol = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
End Function

Private Function GetOrCreateMergeSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateMergeSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateMergeSheet Is Nothing Then
        Set GetOrCreateMergeSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateMergeSheet.Name = sheetName
    End If
End Function
