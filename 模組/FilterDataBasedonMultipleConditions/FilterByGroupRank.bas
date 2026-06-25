Option Explicit
Attribute VB_Name = "FilterByGroupRank"
'*************************************************************************************
'模組名稱: FilterByGroupRank
'功能說明: 依分組排名篩選資料，例如每個部門中銷售額最高的前三名
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestFilterByGroupRank()
    Call FilterByGroupAndRank
End Sub

' 依分組排名篩選資料
Sub FilterByGroupAndRank()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim rng As Range
    Dim arr() As Variant
    Dim departments As Object
    Dim dept As Variant
    Dim tmpRow As Long
    Dim maxRanks As Variant
    Dim resultRow As Long
    Dim currentDept As String

    Set ws = GetOrCreateWorksheet("分組排名資料")
    ws.Cells.Clear

    ' 建立範例銷售資料
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "業務員"
    ws.Range("C1").Value = "銷售額"

    ws.Range("A2").Value = "業務部": ws.Range("B2").Value = "王先生": ws.Range("C2").Value = 120000
    ws.Range("A3").Value = "業務部": ws.Range("B3").Value = "李先生": ws.Range("C3").Value = 95000
    ws.Range("A4").Value = "業務部": ws.Range("B4").Value = "張小姐": ws.Range("C4").Value = 88000
    ws.Range("A5").Value = "業務部": ws.Range("B5").Value = "陳先生": ws.Range("C5").Value = 76000
    ws.Range("A6").Value = "行銷部": ws.Range("B6").Value = "林小姐": ws.Range("C6").Value = 150000
    ws.Range("A7").Value = "行銷部": ws.Range("B7").Value = "黃先生": ws.Range("C7").Value = 135000
    ws.Range("A8").Value = "行銷部": ws.Range("B8").Value = "趙小姐": ws.Range("C8").Value = 110000
    ws.Range("A9").Value = "行銷部": ws.Range("B9").Value = "吳先生": ws.Range("C9").Value = 98000
    ws.Range("A10").Value = "研發部": ws.Range("B10").Value = "周先生": ws.Range("C10").Value = 180000
    ws.Range("A11").Value = "研發部": ws.Range("B11").Value = "蔡小姐": ws.Range("C11").Value = 160000
    ws.Range("A12").Value = "研發部": ws.Range("B12").Value = "劉先生": ws.Range("C12").Value = 145000
    ws.Range("A13").Value = "研發部": ws.Range("B13").Value = "許先生": ws.Range("C13").Value = 130000

    ' 收集所有部門
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set departments = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        currentDept = ws.Cells(i, 1).Value
        If Not departments.Exists(currentDept) Then
            departments.Add currentDept, True
        End If
    Next i

    ' 要求輸入各部門取前幾名
    maxRanks = InputBox("請輸入各部門要篩選的前 N 名：", "前N名", "3")
    If maxRanks = "" Then Exit Sub
    If Not IsNumeric(maxRanks) Then
        MsgBox "請輸入有效的數字。", vbExclamation, "提示"
        Exit Sub
    End If
    maxRanks = CLng(maxRanks)

    ' 建立結果工作表
    Set wsResult = GetOrCreateWorksheet("各部門排名篩選結果")
    wsResult.Cells.Clear
    wsResult.Range("A1").Value = "部門"
    wsResult.Range("B1").Value = "業務員"
    wsResult.Range("C1").Value = "銷售額"
    wsResult.Range("D1").Value = "排名"
    resultRow = 2

    ' 對每個部門進行排序與取前N名
    For Each dept In departments.Keys
        ' 將該部門資料讀入陣列
        ReDim arr(1 To lastRow, 1 To 3)
        tmpRow = 0
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = dept Then
                tmpRow = tmpRow + 1
                arr(tmpRow, 1) = ws.Cells(i, 1).Value
                arr(tmpRow, 2) = ws.Cells(i, 2).Value
                arr(tmpRow, 3) = ws.Cells(i, 3).Value
            End If
        Next i

        ' 依銷售額由大到小排序（簡單氣泡排序）
        Dim k As Long, m As Long
        Dim tmp As Variant
        For k = 1 To tmpRow - 1
            For m = k + 1 To tmpRow
                If arr(m, 3) > arr(k, 3) Then
                    tmp = arr(k, 3): arr(k, 3) = arr(m, 3): arr(m, 3) = tmp
                    Dim tmpName As String, tmpDept As String
                    tmpDept = arr(k, 1): arr(k, 1) = arr(m, 1): arr(m, 1) = tmpDept
                    tmpName = arr(k, 2): arr(k, 2) = arr(m, 2): arr(m, 2) = tmpName
                End If
            Next m
        Next k

        ' 輸出前N名
        Dim limit As Long
        limit = maxRanks
        If limit > tmpRow Then limit = tmpRow
        For k = 1 To limit
            wsResult.Cells(resultRow, 1).Value = arr(k, 1)
            wsResult.Cells(resultRow, 2).Value = arr(k, 2)
            wsResult.Cells(resultRow, 3).Value = arr(k, 3)
            wsResult.Cells(resultRow, 4).Value = k
            resultRow = resultRow + 1
        Next k
    Next dept

    wsResult.Columns.AutoFit

    MsgBox "依部門排名篩選完成，結果已寫入「各部門排名篩選結果」工作表。" & vbCrLf & _
           "共 " & departments.Count & " 個部門，各取前 " & maxRanks & " 名。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
