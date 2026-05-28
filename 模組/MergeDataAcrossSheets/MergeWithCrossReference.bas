Attribute VB_Name = "MergeWithCrossReference"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithCrossReference
'功能說明: 以交叉對照方式合併多張工作表，依據共同鍵值欄位將各表欄位橫向合併到摘要表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestMergeWithCrossReference()
    Call CreateCrossRefDemo(ThisWorkbook)
    Call MergeWithCrossReference(ThisWorkbook, "交叉對照摘要", 1)
End Sub

Private Sub CreateCrossRefDemo(ByVal wb As Workbook)
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    On Error Resume Next
    Set ws1 = wb.Worksheets("員工基本資料")
    Set ws2 = wb.Worksheets("員工薪資資料")
    On Error GoTo 0
    If ws1 Is Nothing Then
        Set ws1 = wb.Worksheets.Add
        ws1.Name = "員工基本資料"
    End If
    ws1.Cells.Clear
    If ws2 Is Nothing Then
        Set ws2 = wb.Worksheets.Add(After:=ws1)
        ws2.Name = "員工薪資資料"
    End If
    ws2.Cells.Clear
    ws1.Range("A1:C1").Value = Array("員工編號", "姓名", "部門")
    ws1.Range("A2:C2").Value = Array("E001", "張小明", "業務部")
    ws1.Range("A3:C3").Value = Array("E002", "李美玲", "財務部")
    ws1.Range("A4:C4").Value = Array("E003", "王大偉", "工程部")
    ws1.Range("A5:C5").Value = Array("E004", "陳佳芳", "人資部")
    ws1.Columns("A:C").AutoFit
    ws2.Range("A1:C1").Value = Array("員工編號", "底薪", "績效獎金")
    ws2.Range("A2:C2").Value = Array("E001", 45000, 8000)
    ws2.Range("A3:C3").Value = Array("E002", 52000, 5000)
    ws2.Range("A4:C4").Value = Array("E003", 60000, 12000)
    ws2.Range("A5:C5").Value = Array("E004", 42000, 3000)
    ws2.Columns("A:C").AutoFit
End Sub

Sub MergeWithCrossReference(ByVal wb As Workbook, _
                             ByVal summaryName As String, _
                             ByVal keyColIndex As Integer)
    Dim summaryWs   As Worksheet
    Dim ws          As Worksheet
    Dim keys        As Object
    Dim allHeaders  As Object
    Dim i           As Long
    Dim j           As Integer
    Dim keyVal      As String
    Dim hdr         As String
    Dim lastRow     As Long
    Dim lastCol     As Integer
    Dim outRow      As Long
    Dim outCol      As Integer

    Set keys = CreateObject("Scripting.Dictionary")
    Set allHeaders = CreateObject("Scripting.Dictionary")

    For Each ws In wb.Worksheets
        If ws.Name <> summaryName Then
            lastRow = ws.Cells(ws.Rows.Count, keyColIndex).End(xlUp).Row
            lastCol = ws.UsedRange.Columns.Count
            For i = 2 To lastRow
                keyVal = CStr(ws.Cells(i, keyColIndex).Value)
                If Len(keyVal) > 0 And Not keys.Exists(keyVal) Then
                    keys.Add keyVal, keyVal
                End If
            Next i
            For j = 1 To lastCol
                hdr = ws.Name & "|" & CStr(ws.Cells(1, j).Value)
                If j <> keyColIndex And Not allHeaders.Exists(hdr) Then
                    allHeaders.Add hdr, allHeaders.Count + 1
                End If
            Next j
        End If
    Next ws

    On Error Resume Next
    Set summaryWs = wb.Worksheets(summaryName)
    On Error GoTo 0
    If summaryWs Is Nothing Then
        Set summaryWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        summaryWs.Name = summaryName
    End If
    summaryWs.Cells.Clear

    summaryWs.Cells(1, 1).Value = wb.Worksheets(1).Cells(1, keyColIndex).Value
    outCol = 2
    Dim hKey As Variant
    For Each hKey In allHeaders.Keys
        summaryWs.Cells(1, outCol).Value = Split(CStr(hKey), "|")(1)
        outCol = outCol + 1
    Next hKey

    With summaryWs.Range(summaryWs.Cells(1, 1), summaryWs.Cells(1, outCol - 1))
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With

    outRow = 2
    Dim kKey As Variant
    For Each kKey In keys.Keys
        keyVal = CStr(kKey)
        summaryWs.Cells(outRow, 1).Value = keyVal
        For Each ws In wb.Worksheets
            If ws.Name <> summaryName Then
                lastRow = ws.Cells(ws.Rows.Count, keyColIndex).End(xlUp).Row
                lastCol = ws.UsedRange.Columns.Count
                For i = 2 To lastRow
                    If CStr(ws.Cells(i, keyColIndex).Value) = keyVal Then
                        For j = 1 To lastCol
                            If j <> keyColIndex Then
                                hdr = ws.Name & "|" & CStr(ws.Cells(1, j).Value)
                                If allHeaders.Exists(hdr) Then
                                    summaryWs.Cells(outRow, allHeaders(hdr) + 1).Value = _
                                        ws.Cells(i, j).Value
                                End If
                            End If
                        Next j
                        Exit For
                    End If
                Next i
            End If
        Next ws
        outRow = outRow + 1
    Next kKey

    summaryWs.Columns.AutoFit
    summaryWs.Activate
    MsgBox "交叉對照合併完成！共 " & keys.Count & " 筆記錄。", vbInformation, "完成"
End Sub
