Option Explicit
Attribute VB_Name = "CompareByRowCount"
'*************************************************************************************
'模組名稱: 以列數比較資料差異
'功能說明: 比較兩張工作表的列數是否一致，並列出列數差異及首欄關鍵值的差集
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub CompareByRowCount()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Set wb = ThisWorkbook

    If wb.Worksheets.Count < 2 Then
        MsgBox "需要至少 2 張工作表才能比較。", vbExclamation, "提示"
        Exit Sub
    End If

    Dim ws1Name As String
    Dim ws2Name As String
    ws1Name = InputBox("請輸入第一張工作表名稱：", "工作表1", wb.Worksheets(1).Name)
    If ws1Name = "" Then Exit Sub

    ws2Name = InputBox("請輸入第二張工作表名稱：", "工作表2", wb.Worksheets(2).Name)
    If ws2Name = "" Then Exit Sub

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    On Error Resume Next
    Set ws1 = wb.Worksheets(ws1Name)
    Set ws2 = wb.Worksheets(ws2Name)
    On Error GoTo ErrorHandler

    If ws1 Is Nothing Or ws2 Is Nothing Then
        MsgBox "找不到指定的工作表，請確認名稱是否正確。", vbExclamation, "錯誤"
        Exit Sub
    End If

    Dim lastRow1 As Long
    Dim lastRow2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row

    ' 扣除標題列
    Dim dataCount1 As Long
    Dim dataCount2 As Long
    dataCount1 = lastRow1 - 1
    dataCount2 = lastRow2 - 1

    Dim report As String
    report = "===== 列數比較報告 =====" & vbCrLf
    report = report & ws1Name & "：" & dataCount1 & " 列資料" & vbCrLf
    report = report & ws2Name & "：" & dataCount2 & " 列資料" & vbCrLf

    If dataCount1 = dataCount2 Then
        report = report & vbCrLf & "兩表列數相同。" & vbCrLf
    Else
        report = report & vbCrLf & "兩表列數不同，差異：" & Abs(dataCount1 - dataCount2) & " 列" & vbCrLf
    End If

    ' 比較首欄關鍵值差集
    Dim keys1 As Object
    Dim keys2 As Object
    Set keys1 = CreateObject("Scripting.Dictionary")
    Set keys2 = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim k1 As String
    Dim k2 As String
    For i = 2 To lastRow1
        k1 = Trim(CStr(ws1.Cells(i, 1).Value))
        If k1 <> "" Then keys1(k1) = 1
    Next i

    For i = 2 To lastRow2
        k2 = Trim(CStr(ws2.Cells(i, 1).Value))
        If k2 <> "" Then keys2(k2) = 1
    Next i

    ' 在 ws1 有但 ws2 沒有
    Dim onlyIn1 As String
    Dim key As Variant
    For Each key In keys1.Keys
        If Not keys2.Exists(key) Then
            onlyIn1 = onlyIn1 & "  " & key & vbCrLf
        End If
    Next key

    ' 在 ws2 有但 ws1 沒有
    Dim onlyIn2 As String
    For Each key In keys2.Keys
        If Not keys1.Exists(key) Then
            onlyIn2 = onlyIn2 & "  " & key & vbCrLf
        End If
    Next key

    If onlyIn1 <> "" Then
        report = report & vbCrLf & "僅在「" & ws1Name & "」的首欄值：" & vbCrLf & onlyIn1
    End If
    If onlyIn2 <> "" Then
        report = report & vbCrLf & "僅在「" & ws2Name & "」的首欄值：" & vbCrLf & onlyIn2
    End If

    MsgBox report, vbInformation, "比較結果"
    Exit Sub

ErrorHandler:
    MsgBox "比較時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
