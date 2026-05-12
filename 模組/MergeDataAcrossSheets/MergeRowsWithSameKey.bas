Option Explicit
Attribute VB_Name = "MergeRowsWithSameKey"
'*************************************************************************************
'模組名稱: MergeRowsWithSameKey
'功能說明: 跨工作表依相同鍵值欄位合併對應列，去除重複後輸出至彙整工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub MergeRowsWithSameKey()
    Dim wbSrc As Workbook
    Dim ws As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim outRow As Long
    Dim keyCol As Integer
    Dim i As Long
    Dim keyVal As String
    Dim dict As Object
    Dim headerWritten As Boolean

    On Error GoTo ErrHandler

    Set wbSrc = ThisWorkbook
    keyCol = 1
    headerWritten = False
    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set wsOut = wbSrc.Sheets("KeyMergeResult")
    On Error GoTo ErrHandler

    If wsOut Is Nothing Then
        Set wsOut = wbSrc.Sheets.Add(After:=wbSrc.Sheets(wbSrc.Sheets.Count))
        wsOut.Name = "KeyMergeResult"
    Else
        wsOut.Cells.Clear
    End If

    outRow = 2

    For Each ws In wbSrc.Worksheets
        If ws.Name <> "KeyMergeResult" Then
            lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row

            If Not headerWritten And lastRow >= 1 Then
                ws.Rows(1).Copy Destination:=wsOut.Rows(1)
                headerWritten = True
            End If

            For i = 2 To lastRow
                keyVal = Trim(CStr(ws.Cells(i, keyCol).Value))
                If keyVal <> "" Then
                    If Not dict.Exists(keyVal) Then
                        ws.Rows(i).Copy Destination:=wsOut.Rows(outRow)
                        dict.Add keyVal, outRow
                        outRow = outRow + 1
                    End If
                End If
            Next i
        End If
    Next ws

    wsOut.Columns.AutoFit
    MsgBox "已依相同鍵值合併列資料，共匯出 " & dict.Count & " 筆唯一記錄。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "依鍵值合併列失敗"
End Sub