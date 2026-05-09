Attribute VB_Name = "FilterByDateRangeWithDropdown"
Option Explicit

'************************************************************************************
' 模組名稱: FilterByDateRangeWithDropdown
' 功能說明: 使用 AutoFilter 依「月份」或「季度」下拉選項篩選日期欄位
'           提供快速按月／按季篩選的工具函式
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 入口：依月份篩選（只保留指定年月的資料列）
Public Sub FilterByMonthExample()
    On Error GoTo ErrHandler

    Dim ws          As Worksheet
    Dim targetYear  As Integer
    Dim targetMonth As Integer

    Set ws = GetOrCreateWsDate(ThisWorkbook, "月份篩選範例")
    Call FillMonthlyData(ws)

    targetYear  = 2026
    targetMonth = 3  ' 3月

    Call ApplyMonthFilter(ws, 1, targetYear, targetMonth)

    MsgBox "已篩選 " & targetYear & " 年 " & targetMonth & " 月的資料。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 入口：依季度篩選（Q1=1~3月, Q2=4~6月, ...）
Public Sub FilterByQuarterExample()
    On Error GoTo ErrHandler

    Dim ws        As Worksheet
    Dim targetQtr As Integer

    Set ws = GetOrCreateWsDate(ThisWorkbook, "季度篩選範例")
    Call FillMonthlyData(ws)

    targetQtr = 1  ' 第一季 (1~3月)

    Call ApplyQuarterFilter(ws, 1, 2026, targetQtr)

    MsgBox "已篩選 2026 年第 " & targetQtr & " 季資料。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 依年月篩選（在 Field 欄套用開始與結束日期範圍）
Private Sub ApplyMonthFilter(ByVal ws As Worksheet, ByVal dateField As Integer, _
                              ByVal yr As Integer, ByVal mo As Integer)
    Dim dtStart As Date
    Dim dtEnd   As Date

    dtStart = DateSerial(yr, mo, 1)
    dtEnd   = DateSerial(yr, mo + 1, 1) - 1

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ws.Range("A1").CurrentRegion.AutoFilter _
        Field:=dateField, _
        Criteria1:=">=" & CLng(dtStart), _
        Operator:=xlAnd, _
        Criteria2:="<=" & CLng(dtEnd)

    ws.Columns("A:C").AutoFit
End Sub

' 依季度篩選
Private Sub ApplyQuarterFilter(ByVal ws As Worksheet, ByVal dateField As Integer, _
                                ByVal yr As Integer, ByVal qtr As Integer)
    Dim startMonth As Integer
    Dim dtStart    As Date
    Dim dtEnd      As Date

    startMonth = (qtr - 1) * 3 + 1
    dtStart = DateSerial(yr, startMonth, 1)
    dtEnd   = DateSerial(yr, startMonth + 3, 1) - 1

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ws.Range("A1").CurrentRegion.AutoFilter _
        Field:=dateField, _
        Criteria1:=">=" & CLng(dtStart), _
        Operator:=xlAnd, _
        Criteria2:="<=" & CLng(dtEnd)

    ws.Columns("A:C").AutoFit
End Sub

' 填入逐月銷售測試資料
Private Sub FillMonthlyData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("日期", "品項", "銷售額")
    ws.Range("A2:C2").Value  = Array(DateSerial(2026, 1, 5),  "商品A", 32000)
    ws.Range("A3:C3").Value  = Array(DateSerial(2026, 1, 20), "商品B", 18000)
    ws.Range("A4:C4").Value  = Array(DateSerial(2026, 2, 10), "商品A", 41000)
    ws.Range("A5:C5").Value  = Array(DateSerial(2026, 2, 25), "商品C", 27000)
    ws.Range("A6:C6").Value  = Array(DateSerial(2026, 3, 8),  "商品B", 55000)
    ws.Range("A7:C7").Value  = Array(DateSerial(2026, 3, 22), "商品A", 38000)
    ws.Range("A8:C8").Value  = Array(DateSerial(2026, 4, 3),  "商品C", 62000)
    ws.Range("A9:C9").Value  = Array(DateSerial(2026, 4, 18), "商品A", 29000)
    ws.Range("A2:A9").NumberFormat = "yyyy/mm/dd"
    ws.Range("C2:C9").NumberFormat = "#,##0"
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsDate(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWsDate = ws
End Function