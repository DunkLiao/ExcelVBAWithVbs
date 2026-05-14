Attribute VB_Name = "FilterAndHighlight"
Option Explicit
'*************************************************************************************
'模組名稱: FilterAndHighlight
'功能說明: 依多重條件篩選資料列，並以背景色高亮標示符合條件的列（不移除其他列）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestFilterAndHighlight()
    Call CreateHighlightTestData
    Call FilterAndHighlightRows(ActiveSheet, 2, "業務部", 3, 80000)
End Sub

' 建立測試資料
Private Sub CreateHighlightTestData()
    Dim ws As Worksheet
    Set ws = GetOrCreateHighlightWs(ThisWorkbook, "高亮篩選範例")
    ws.Cells.Clear
    ws.Activate

    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "業績"
    ws.Range("D1").Value = "評等"
    ws.Range("A1:D1").Font.Bold = True

    Dim data(1 To 10, 1 To 4) As Variant
    data(1, 1) = "王小明" : data(1, 2) = "業務部" : data(1, 3) = 92000 : data(1, 4) = "優"
    data(2, 1) = "李大華" : data(2, 2) = "行銷部" : data(2, 3) = 75000 : data(2, 4) = "良"
    data(3, 1) = "陳美玲" : data(3, 2) = "業務部" : data(3, 3) = 85000 : data(3, 4) = "優"
    data(4, 1) = "林俊傑" : data(4, 2) = "研發部" : data(4, 3) = 68000 : data(4, 4) = "可"
    data(5, 1) = "張志遠" : data(5, 2) = "業務部" : data(5, 3) = 110000 : data(5, 4) = "優"
    data(6, 1) = "吳雅婷" : data(6, 2) = "行銷部" : data(6, 3) = 88000 : data(6, 4) = "優"
    data(7, 1) = "黃建宏" : data(7, 2) = "業務部" : data(7, 3) = 62000 : data(7, 4) = "可"
    data(8, 1) = "劉佳欣" : data(8, 2) = "客服部" : data(8, 3) = 55000 : data(8, 4) = "可"
    data(9, 1) = "蔡文峰" : data(9, 2) = "業務部" : data(9, 3) = 98000 : data(9, 4) = "優"
    data(10, 1) = "許志明" : data(10, 2) = "研發部" : data(10, 3) = 72000 : data(10, 4) = "良"

    Dim i As Long
    For i = 1 To 10
        ws.Cells(i + 1, 1).Value = data(i, 1)
        ws.Cells(i + 1, 2).Value = data(i, 2)
        ws.Cells(i + 1, 3).Value = data(i, 3)
        ws.Cells(i + 1, 4).Value = data(i, 4)
    Next i

    ws.Columns("A:D").AutoFit
End Sub

' 依條件篩選並高亮標示符合列
' ws           : 目標工作表
' textCol      : 文字比對欄索引
' textCriteria : 文字條件（含此字串即符合）
' numCol       : 數值比對欄索引
' minValue     : 數值下限（大於等於時符合）
Sub FilterAndHighlightRows(ByVal ws As Worksheet, ByVal textCol As Long, _
    ByVal textCriteria As String, ByVal numCol As Long, ByVal minValue As Double)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = xlNone

    Dim matchCount As Long
    matchCount = 0
    Dim highlightColor As Long
    highlightColor = RGB(255, 255, 153)

    Dim i As Long
    For i = 2 To lastRow
        Dim textMatch As Boolean
        Dim numMatch As Boolean

        textMatch = (InStr(1, CStr(ws.Cells(i, textCol).Value), textCriteria, vbTextCompare) > 0)
        numMatch = (IsNumeric(ws.Cells(i, numCol).Value) And _
            CDbl(ws.Cells(i, numCol).Value) >= minValue)

        If textMatch And numMatch Then
            ws.Rows(i).Interior.Color = highlightColor
            matchCount = matchCount + 1
        End If
    Next i

    MsgBox "高亮篩選完成！" & vbCrLf & _
        "條件：" & ws.Cells(1, textCol).Value & " 包含 " & textCriteria & _
        "  且  " & ws.Cells(1, numCol).Value & " >= " & minValue & vbCrLf & _
        "符合筆數：" & matchCount & " 筆", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "篩選高亮時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 清除所有高亮標示
Sub ClearAllHighlights()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = xlNone
    End If

    MsgBox "已清除所有高亮標示。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "清除高亮時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateHighlightWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateHighlightWs = ws
End Function