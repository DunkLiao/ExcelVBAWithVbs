Attribute VB_Name = "CompareByHashValue"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByHashValue
'功能說明: 以雜湊值比對兩個工作表列資料差異，快速找出新增、刪除、修改的列
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestCompareByHashValue()
    Call CompareSheetsByHash
End Sub

' 以雜湊值比對兩個工作表
Sub CompareSheetsByHash()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim wsResult As Worksheet
    Dim lngRow1 As Long
    Dim lngRow2 As Long
    Dim lngResultRow As Long
    Dim i As Long
    Dim j As Long
    Dim sHash1 As String
    Dim sHash2 As String
    Dim blnFound As Boolean

    On Error GoTo ErrHandler

    If ThisWorkbook.Worksheets.Count < 2 Then
        MsgBox "需要至少兩個工作表才能進行比對。", vbExclamation
        Exit Sub
    End If

    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    Set wsResult = GetOrCreateHashResultSheet(ThisWorkbook, "雜湊比對結果")

    wsResult.Range("A1").Value = "狀態"
    wsResult.Range("B1").Value = "來源工作表"
    wsResult.Range("C1").Value = "列號"
    wsResult.Range("D1").Value = "雜湊值"
    wsResult.Range("E1").Value = "列內容摘要"
    wsResult.Range("A1:E1").Font.Bold = True
    lngResultRow = 2

    lngRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lngRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 2 To lngRow1
        sHash1 = ComputeRowHash(ws1, i)
        blnFound = False
        For j = 2 To lngRow2
            sHash2 = ComputeRowHash(ws2, j)
            If sHash1 = sHash2 Then
                blnFound = True
                Exit For
            End If
        Next j
        If Not blnFound Then
            wsResult.Cells(lngResultRow, 1).Value = "僅在表1（可能已刪除）"
            wsResult.Cells(lngResultRow, 1).Font.Color = RGB(255, 0, 0)
            wsResult.Cells(lngResultRow, 2).Value = ws1.Name
            wsResult.Cells(lngResultRow, 3).Value = i
            wsResult.Cells(lngResultRow, 4).Value = sHash1
            wsResult.Cells(lngResultRow, 5).Value = GetRowSummary(ws1, i)
            lngResultRow = lngResultRow + 1
        End If
    Next i

    For j = 2 To lngRow2
        sHash2 = ComputeRowHash(ws2, j)
        blnFound = False
        For i = 2 To lngRow1
            sHash1 = ComputeRowHash(ws1, i)
            If sHash2 = sHash1 Then
                blnFound = True
                Exit For
            End If
        Next i
        If Not blnFound Then
            wsResult.Cells(lngResultRow, 1).Value = "僅在表2（可能為新增）"
            wsResult.Cells(lngResultRow, 1).Font.Color = RGB(0, 128, 0)
            wsResult.Cells(lngResultRow, 2).Value = ws2.Name
            wsResult.Cells(lngResultRow, 3).Value = j
            wsResult.Cells(lngResultRow, 4).Value = sHash2
            wsResult.Cells(lngResultRow, 5).Value = GetRowSummary(ws2, j)
            lngResultRow = lngResultRow + 1
        End If
    Next j

    wsResult.Columns("A:E").AutoFit
    wsResult.Activate
    Application.ScreenUpdating = True

    Dim lngDiff As Long
    lngDiff = lngResultRow - 2
    If lngDiff = 0 Then
        MsgBox "兩個工作表的資料完全相同（雜湊值一致）！", vbInformation, "比對完成"
    Else
        MsgBox "比對完成，共發現 " & lngDiff & " 列差異，詳見「雜湊比對結果」工作表。", _
            vbInformation, "比對完成"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 計算列的簡易雜湊值
Private Function ComputeRowHash(ByVal ws As Worksheet, ByVal lngRow As Long) As String
    Dim lngLastCol As Long
    Dim sContent As String
    Dim lngSum As Long
    Dim k As Long
    Dim m As Integer

    lngLastCol = ws.Cells(lngRow, ws.Columns.Count).End(xlToLeft).Column
    sContent = ""
    For k = 1 To lngLastCol
        sContent = sContent & "|" & CStr(ws.Cells(lngRow, k).Value)
    Next k

    lngSum = Len(sContent)
    For m = 1 To Len(sContent)
        lngSum = lngSum + Asc(Mid(sContent, m, 1)) * m
    Next m

    ComputeRowHash = CStr(Len(sContent)) & "_" & CStr(lngSum Mod 99991)
End Function

' 取得列內容摘要（前三欄值）
Private Function GetRowSummary(ByVal ws As Worksheet, ByVal lngRow As Long) As String
    Dim s As String
    s = CStr(ws.Cells(lngRow, 1).Value)
    If ws.Cells(lngRow, 2).Value <> "" Then
        s = s & " | " & CStr(ws.Cells(lngRow, 2).Value)
    End If
    If ws.Cells(lngRow, 3).Value <> "" Then
        s = s & " | " & CStr(ws.Cells(lngRow, 3).Value)
    End If
    GetRowSummary = s
End Function

' 取得或建立結果工作表並清除內容
Private Function GetOrCreateHashResultSheet(ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateHashResultSheet = ws
End Function
