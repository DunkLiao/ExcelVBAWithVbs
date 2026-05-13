Attribute VB_Name = "FilterByNamedRange"
Option Explicit
'*************************************************************************************
'模組名稱: 依具名範圍篩選資料
'功能說明: 從具名範圍讀取篩選清單，對工作表資料執行自動篩選
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub FilterByNamedRange()
    Dim ws As Worksheet
    Dim criteriaRange As Range
    Dim namedRangeName As String
    Dim filterColIdx As Long
    Dim filterColStr As String
    Dim criteriaArr() As String
    Dim i As Long
    Dim cell As Range
    Dim criteriaCount As Long

    namedRangeName = InputBox("請輸入具名範圍名稱（包含篩選清單）：", "具名範圍", "篩選清單")
    If namedRangeName = "" Then Exit Sub

    On Error Resume Next
    Set criteriaRange = ThisWorkbook.Names(namedRangeName).RefersToRange
    On Error GoTo 0

    If criteriaRange Is Nothing Then
        MsgBox "找不到具名範圍：" & namedRangeName & vbCrLf & _
               "請先在工作表中定義此具名範圍。", vbExclamation, "錯誤"
        Exit Sub
    End If

    filterColStr = InputBox("請輸入要篩選的欄號（例如：1）：", "欄號", "1")
    If filterColStr = "" Then Exit Sub
    If Not IsNumeric(filterColStr) Then
        MsgBox "請輸入有效的欄號。", vbExclamation, "錯誤"
        Exit Sub
    End If
    filterColIdx = CLng(filterColStr)

    criteriaCount = 0
    For Each cell In criteriaRange
        If CStr(cell.Value) <> "" Then
            criteriaCount = criteriaCount + 1
        End If
    Next cell

    If criteriaCount = 0 Then
        MsgBox "具名範圍「" & namedRangeName & "」中沒有篩選條件。", vbExclamation, "提示"
        Exit Sub
    End If

    ReDim criteriaArr(0 To criteriaCount - 1)
    i = 0
    For Each cell In criteriaRange
        If CStr(cell.Value) <> "" Then
            criteriaArr(i) = CStr(cell.Value)
            i = i + 1
        End If
    Next cell

    Set ws = ActiveSheet

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    ws.UsedRange.AutoFilter Field:=filterColIdx, _
        Criteria1:=criteriaArr, Operator:=xlFilterValues

    MsgBox "已依具名範圍「" & namedRangeName & "」套用篩選，" & _
           "共 " & criteriaCount & " 個條件。", vbInformation, "完成"
End Sub

Sub CreateFilterByNamedRangeDemo()
    Dim ws As Worksheet
    Dim wsCriteria As Worksheet
    Dim demoName As String

    demoName = "具名範圍篩選範例"

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(demoName).Delete
    ThisWorkbook.Worksheets("篩選條件").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = demoName

    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "業績"

    ws.Range("A2").Value = "業務部" : ws.Range("B2").Value = "王小明" : ws.Range("C2").Value = 80000
    ws.Range("A3").Value = "行政部" : ws.Range("B3").Value = "李美玲" : ws.Range("C3").Value = 45000
    ws.Range("A4").Value = "業務部" : ws.Range("B4").Value = "張大為" : ws.Range("C4").Value = 95000
    ws.Range("A5").Value = "研發部" : ws.Range("B5").Value = "陳志明" : ws.Range("C5").Value = 72000
    ws.Range("A6").Value = "行政部" : ws.Range("B6").Value = "林淑華" : ws.Range("C6").Value = 38000
    ws.Range("A7").Value = "研發部" : ws.Range("B7").Value = "黃建國" : ws.Range("C7").Value = 68000

    Set wsCriteria = ThisWorkbook.Worksheets.Add
    wsCriteria.Name = "篩選條件"

    wsCriteria.Range("A1").Value = "篩選清單"
    wsCriteria.Range("A2").Value = "業務部"
    wsCriteria.Range("A3").Value = "研發部"

    On Error Resume Next
    ThisWorkbook.Names("篩選清單").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="篩選清單", RefersTo:=wsCriteria.Range("A2:A3")

    ws.Columns("A:C").AutoFit
    ws.Range("A1:C1").Font.Bold = True

    MsgBox "示範資料已建立！" & vbCrLf & _
           "請切換到「" & demoName & "」工作表後執行 FilterByNamedRange。" & vbCrLf & _
           "具名範圍名稱：篩選清單，篩選欄號：1", vbInformation, "提示"
End Sub
