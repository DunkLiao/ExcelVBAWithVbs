Option Explicit
Attribute VB_Name = "CompareAndExportReport"
'*************************************************************************************
'模組名稱: CompareAndExportReport
'功能說明: 比較兩個工作表的資料差異並將差異報告匯出至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Sub CompareAndExportReport()
    Dim ws1         As Worksheet
    Dim ws2         As Worksheet
    Dim wsReport    As Worksheet
    Dim sheet1Name  As String
    Dim sheet2Name  As String
    Dim lastRow1    As Long
    Dim lastRow2    As Long
    Dim lastCol     As Long
    Dim i           As Long
    Dim j           As Long
    Dim reportRow   As Long
    Dim maxRow      As Long
    Dim val1        As String
    Dim val2        As String
    Dim cellVal     As Double

    sheet1Name = InputBox("請輸入第一個工作表名稱（舊資料）：", "比較工作表")
    If sheet1Name = "" Then Exit Sub
    sheet2Name = InputBox("請輸入第二個工作表名稱（新資料）：", "比較工作表")
    If sheet2Name = "" Then Exit Sub

    On Error GoTo ErrHandler
    Set ws1 = ThisWorkbook.Sheets(sheet1Name)
    Set ws2 = ThisWorkbook.Sheets(sheet2Name)
    On Error GoTo 0

    ' 建立報告工作表
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("差異報告").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsReport = ThisWorkbook.Sheets.Add( _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsReport.Name = "差異報告"

    ' 寫入標題
    wsReport.Cells(1, 1).Value = "列號"
    wsReport.Cells(1, 2).Value = "欄號"
    wsReport.Cells(1, 3).Value = "欄位名稱"
    wsReport.Cells(1, 4).Value = "舊值"
    wsReport.Cells(1, 5).Value = "新值"
    wsReport.Rows(1).Font.Bold = True
    reportRow = 2

    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    maxRow = lastRow1
    If lastRow2 > maxRow Then maxRow = lastRow2

    ' 逐列逐欄比較
    For i = 2 To maxRow
        For j = 1 To lastCol
            val1 = ""
            val2 = ""
            If i <= lastRow1 Then val1 = CStr(ws1.Cells(i, j).Value)
            If i <= lastRow2 Then val2 = CStr(ws2.Cells(i, j).Value)
            If val1 <> val2 Then
                wsReport.Cells(reportRow, 1).Value = i
                wsReport.Cells(reportRow, 2).Value = j
                wsReport.Cells(reportRow, 3).Value = ws1.Cells(1, j).Value
                wsReport.Cells(reportRow, 4).Value = val1
                wsReport.Cells(reportRow, 5).Value = val2
                wsReport.Rows(reportRow).Interior.Color = RGB(255, 235, 156)
                reportRow = reportRow + 1
            End If
        Next j
    Next i

    wsReport.Columns("A:E").AutoFit
    MsgBox "比較完成，共發現 " & (reportRow - 2) & " 處差異。" & vbCrLf & _
           "報告已輸出至差異報告工作表。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "找不到工作表：" & Err.Description, vbCritical, "錯誤"
End Sub
