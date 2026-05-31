Attribute VB_Name = "FilterByCompositeCondition"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByCompositeCondition
'功能說明: 依複合條件（AND+OR組合）篩選資料列並複製至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestFilterByCompositeCondition()
    Dim ws As Worksheet
    Set ws = GetOrCreateFilterSheet(ThisWorkbook, "複合條件篩選來源")
    Call FillFilterSampleData(ws)
    Call FilterByCompositeCondition(ws, "複合條件篩選結果")
    MsgBox "複合條件篩選完成！", vbInformation, "完成"
End Sub

Sub FilterByCompositeCondition(ByVal srcWs As Worksheet, ByVal destSheetName As String)
    Dim destWs   As Worksheet
    Dim lastRow  As Long
    Dim i        As Long
    Dim destRow  As Long
    Dim dept     As String
    Dim amount   As Double
    Dim status   As String
    Dim match1   As Boolean
    Dim match2   As Boolean

    On Error Resume Next
    Set destWs = ThisWorkbook.Worksheets(destSheetName)
    On Error GoTo 0
    If destWs Is Nothing Then
        Set destWs = ThisWorkbook.Worksheets.Add
        destWs.Name = destSheetName
    End If
    destWs.Cells.Clear

    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "來源資料不足。", vbExclamation, "警告"
        Exit Sub
    End If

    srcWs.Rows(1).Copy Destination:=destWs.Rows(1)
    destRow = 2

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        dept   = Trim(CStr(srcWs.Cells(i, 1).Value))
        amount = 0
        On Error Resume Next
        amount = CDbl(srcWs.Cells(i, 3).Value)
        On Error GoTo 0
        status = Trim(CStr(srcWs.Cells(i, 4).Value))

        match1 = (dept = "業務部") And (amount > 50000)
        match2 = (status = "已完成") And (amount > 80000)

        If match1 Or match2 Then
            srcWs.Rows(i).Copy Destination:=destWs.Rows(destRow)
            destRow = destRow + 1
        End If
    Next i

    Application.ScreenUpdating = True
    destWs.Columns.AutoFit

    MsgBox "複合條件篩選完成！共篩選出 " & destRow - 2 & " 列資料。", _
           vbInformation, "完成"
End Sub

Private Sub FillFilterSampleData(ByVal ws As Worksheet)
    Dim dataArr As Variant
    Dim i       As Integer

    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("部門", "月份", "金額", "狀態")

    dataArr = Array( _
        Array("業務部", "一月", 45000, "已完成"), _
        Array("業務部", "一月", 62000, "已完成"), _
        Array("行銷部", "二月", 85000, "已完成"), _
        Array("業務部", "二月", 55000, "進行中"), _
        Array("行銷部", "二月", 28000, "取消"), _
        Array("業務部", "三月", 91000, "已完成"), _
        Array("行銷部", "三月", 41000, "進行中"), _
        Array("業務部", "三月", 76000, "已完成"))

    For i = 0 To UBound(dataArr)
        ws.Cells(i + 2, 1).Value = dataArr(i)(0)
        ws.Cells(i + 2, 2).Value = dataArr(i)(1)
        ws.Cells(i + 2, 3).Value = dataArr(i)(2)
        ws.Cells(i + 2, 4).Value = dataArr(i)(3)
    Next i

    ws.Columns("A:D").AutoFit
End Sub

Private Function GetOrCreateFilterSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateFilterSheet = ws
End Function
