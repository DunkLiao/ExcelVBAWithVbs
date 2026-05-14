Attribute VB_Name = "MergeWithMatchedHeaders"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithMatchedHeaders
'功能說明: 依欄位名稱比對，將多個工作表資料合併到一個彙總工作表（欄位順序不同也能正確對齊）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestMergeWithMatchedHeaders()
    Call CreateMatchedHeaderTestData
    Call MergeWithMatchedHeaders
End Sub

' 建立測試用資料（欄位順序不同）
Private Sub CreateMatchedHeaderTestData()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim ws1 As Worksheet
    Set ws1 = GetOrCreateHdrWs(wb, "部門A資料")
    ws1.Cells.Clear
    ws1.Range("A1").Value = "姓名"
    ws1.Range("B1").Value = "部門"
    ws1.Range("C1").Value = "業績"
    ws1.Range("A2").Value = "王小明" : ws1.Range("B2").Value = "業務部" : ws1.Range("C2").Value = 85000
    ws1.Range("A3").Value = "李大華" : ws1.Range("B3").Value = "業務部" : ws1.Range("C3").Value = 72000

    Dim ws2 As Worksheet
    Set ws2 = GetOrCreateHdrWs(wb, "部門B資料")
    ws2.Cells.Clear
    ws2.Range("A1").Value = "業績"
    ws2.Range("B1").Value = "姓名"
    ws2.Range("C1").Value = "部門"
    ws2.Range("A2").Value = 91000 : ws2.Range("B2").Value = "陳美玲" : ws2.Range("C2").Value = "行銷部"
    ws2.Range("A3").Value = 68000 : ws2.Range("B3").Value = "林俊傑" : ws2.Range("C3").Value = "行銷部"
End Sub

' 依欄位名稱比對合併所有工作表（排除彙總工作表本身）
Sub MergeWithMatchedHeaders()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim wsResult As Worksheet
    Dim targetSheetName As String
    targetSheetName = "欄位對齊合併"
    Set wsResult = GetOrCreateHdrWs(wb, targetSheetName)
    wsResult.Cells.Clear

    ' 收集所有欄位名稱（聯集）
    Dim allHeaders() As String
    Dim headerCount As Long
    headerCount = 0
    ReDim allHeaders(0)

    Dim ws As Worksheet
    Dim headerCol As Long
    Dim headerName As String
    Dim found As Boolean
    Dim j As Long

    For Each ws In wb.Worksheets
        If ws.Name <> targetSheetName Then
            headerCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            Dim k As Long
            For k = 1 To headerCol
                headerName = CStr(ws.Cells(1, k).Value)
                If headerName <> "" Then
                    found = False
                    For j = 0 To headerCount - 1
                        If allHeaders(j) = headerName Then
                            found = True
                            Exit For
                        End If
                    Next j
                    If Not found Then
                        ReDim Preserve allHeaders(headerCount)
                        allHeaders(headerCount) = headerName
                        headerCount = headerCount + 1
                    End If
                End If
            Next k
        End If
    Next ws

    If headerCount = 0 Then
        MsgBox "找不到任何欄位名稱", vbExclamation, "提示"
        Exit Sub
    End If

    ' 寫入彙總標題
    Dim h As Long
    For h = 0 To headerCount - 1
        wsResult.Cells(1, h + 1).Value = allHeaders(h)
    Next h
    wsResult.Rows(1).Font.Bold = True

    Dim targetRow As Long
    targetRow = 2

    ' 逐工作表複製資料（按欄位名稱對應）
    For Each ws In wb.Worksheets
        If ws.Name <> targetSheetName Then
            Dim srcLastRow As Long
            srcLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            If srcLastRow >= 2 Then
                ' 建立來源欄位名稱→欄號對應
                Dim srcLastCol As Long
                srcLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

                Dim srcColMap(100) As Long
                Dim m As Long
                For m = 0 To headerCount - 1
                    srcColMap(m) = 0
                    Dim n As Long
                    For n = 1 To srcLastCol
                        If CStr(ws.Cells(1, n).Value) = allHeaders(m) Then
                            srcColMap(m) = n
                            Exit For
                        End If
                    Next n
                Next m

                ' 複製每一列資料
                Dim r As Long
                For r = 2 To srcLastRow
                    For m = 0 To headerCount - 1
                        If srcColMap(m) > 0 Then
                            wsResult.Cells(targetRow, m + 1).Value = _
                                ws.Cells(r, srcColMap(m)).Value
                        End If
                    Next m
                    targetRow = targetRow + 1
                Next r
            End If
        End If
    Next ws

    wsResult.UsedRange.Columns.AutoFit
    wsResult.Activate

    MsgBox "依欄位名稱合併完成，共 " & targetRow - 2 & " 列資料。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateHdrWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateHdrWs = ws
End Function
