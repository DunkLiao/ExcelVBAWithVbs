Option Explicit
Attribute VB_Name = "MergeAllVisibleSheets"
'*************************************************************************************
'模組名稱: MergeAllVisibleSheets
'功能說明: 合併活頁簿中所有可見工作表的資料至彙總工作表（跳過隱藏工作表）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Sub MergeAllVisibleSheets()
    Dim wsSummary       As Worksheet
    Dim ws              As Worksheet
    Dim lastRow         As Long
    Dim lastCol         As Long
    Dim nextRow         As Long
    Dim headerCopied    As Boolean

    ' 建立彙總工作表
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("彙總")
    On Error GoTo 0
    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSummary.Name = "彙總"
    Else
        wsSummary.Cells.Clear
    End If

    nextRow = 1
    headerCopied = False

    For Each ws In ThisWorkbook.Sheets
        ' 跳過隱藏工作表及彙總工作表本身
        If ws.Visible = xlSheetVisible And ws.Name <> "彙總" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            If lastRow >= 1 And lastCol >= 1 Then
                If Not headerCopied Then
                    ws.Rows(1).Copy wsSummary.Rows(nextRow)
                    nextRow = nextRow + 1
                    headerCopied = True
                End If

                If lastRow >= 2 Then
                    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy _
                        wsSummary.Cells(nextRow, 1)
                    nextRow = nextRow + (lastRow - 1)
                End If
            End If
        End If
    Next ws

    wsSummary.Columns.AutoFit
    MsgBox "所有可見工作表已合併至彙總工作表，共 " & (nextRow - 2) & " 筆資料。", _
           vbInformation, "完成"
End Sub
