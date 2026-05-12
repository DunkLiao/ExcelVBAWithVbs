Option Explicit
Attribute VB_Name = "MergeSheetsWithSameStructure"
'*************************************************************************************
'模組名稱: MergeSheetsWithSameStructure
'功能說明: 將具有相同欄位結構的多個工作表合併至一張彙總工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub MergeSheetsWithSameStructure()
    Dim wbSrc As Workbook
    Dim wsSummary As Worksheet
    Dim ws As Worksheet
    Dim destRow As Long
    Dim srcLastRow As Long
    Dim headerCopied As Boolean

    On Error GoTo ErrHandler

    Set wbSrc = ThisWorkbook

    On Error Resume Next
    Set wsSummary = wbSrc.Sheets("彙總")
    On Error GoTo ErrHandler

    If wsSummary Is Nothing Then
        Set wsSummary = wbSrc.Sheets.Add(After:=wbSrc.Sheets(wbSrc.Sheets.Count))
        wsSummary.Name = "彙總"
    Else
        wsSummary.Cells.Clear
    End If

    destRow = 1
    headerCopied = False

    For Each ws In wbSrc.Worksheets
        If ws.Name <> "彙總" Then
            srcLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If srcLastRow >= 1 Then
                If Not headerCopied Then
                    ws.Rows(1).Copy Destination:=wsSummary.Rows(destRow)
                    destRow = destRow + 1
                    headerCopied = True
                End If
                If srcLastRow > 1 Then
                    ws.Rows("2:" & srcLastRow).Copy Destination:=wsSummary.Rows(destRow)
                    destRow = destRow + srcLastRow - 1
                End If
            End If
        End If
    Next ws

    wsSummary.Columns.AutoFit
    MsgBox "相同結構工作表已合併至「彙總」工作表，共 " & destRow - 2 & " 筆資料。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "合併工作表失敗"
End Sub