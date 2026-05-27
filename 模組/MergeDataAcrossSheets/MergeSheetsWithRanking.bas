Option Explicit
Attribute VB_Name = "MergeSheetsWithRanking"
'*************************************************************************************
'模組名稱: 跨工作表合併並排名
'功能說明: 將所有工作表第一欄（名稱）與第二欄（數值）合併到彙總表，
'          並在第三欄自動加上排名（由大到小）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub MergeSheetsWithRanking()
    On Error GoTo ErrorHandler

    Dim wbk As Workbook
    Dim wsSummary As Worksheet
    Dim ws As Worksheet
    Dim destRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim summaryName As String

    summaryName = "排名彙總"
    Set wbk = ThisWorkbook

    ' 刪除舊的彙總工作表（若存在）
    Application.DisplayAlerts = False
    On Error Resume Next
    wbk.Worksheets(summaryName).Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    Set wsSummary = wbk.Worksheets.Add(After:=wbk.Worksheets(wbk.Worksheets.Count))
    wsSummary.Name = summaryName

    wsSummary.Range("A1").Value = "來源工作表"
    wsSummary.Range("B1").Value = "名稱"
    wsSummary.Range("C1").Value = "數值"
    wsSummary.Range("D1").Value = "排名"
    destRow = 2

    For Each ws In wbk.Worksheets
        If ws.Name <> summaryName Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            For i = 2 To lastRow
                If Trim(CStr(ws.Cells(i, 1).Value)) <> "" Then
                    wsSummary.Cells(destRow, 1).Value = ws.Name
                    wsSummary.Cells(destRow, 2).Value = ws.Cells(i, 1).Value
                    wsSummary.Cells(destRow, 3).Value = ws.Cells(i, 2).Value
                    destRow = destRow + 1
                End If
            Next i
        End If
    Next ws

    If destRow > 2 Then
        ' 寫入 RANK 公式
        Dim dataEnd As Long
        dataEnd = destRow - 1
        Dim rankRange As String
        rankRange = "$C$2:$C$" & dataEnd

        For i = 2 To dataEnd
            wsSummary.Cells(i, 4).Formula = "=RANK(C" & i & "," & rankRange & ",0)"
        Next i

        ' 依排名排序
        With wsSummary.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsSummary.Range("D2:D" & dataEnd), _
                SortOn:=xlSortOnValues, Order:=xlAscending
            .SetRange wsSummary.Range("A1:D" & dataEnd)
            .Header = xlYes
            .Apply
        End With
    End If

    wsSummary.Columns("A:D").AutoFit
    MsgBox "跨工作表合併並排名完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "合併排名時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
