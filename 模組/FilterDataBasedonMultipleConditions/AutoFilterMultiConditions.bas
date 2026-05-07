Attribute VB_Name = "AutoFilterMultiConditions"
Option Explicit
'*************************************************************************************
'模組名稱: AutoFilterMultiConditions
'功能說明: 示範依據多重條件篩選資料，包含自動篩選及進階篩選
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口
Sub TestAutoFilterMultiConditions()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "篩選資料範例")
    Call FillEmployeeData(ws)
    Call FilterByMultipleConditions(ws, "研發部", 55000)
End Sub

' 依多重條件篩選（部門 + 薪資下限）
Sub FilterByMultipleConditions(ByVal ws As Worksheet, ByVal department As String, _
                                ByVal minSalary As Long)
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    ' 套用自動篩選：A欄=姓名, B欄=部門, C欄=薪資
    ws.Range("A1").AutoFilter
    ws.Range("A1").AutoFilter Field:=2, Criteria1:=department
    ws.Range("A1").AutoFilter Field:=3, Criteria1:=">" & minSalary

    MsgBox "已套用篩選條件：" & vbCrLf & _
           "部門 = " & department & vbCrLf & _
           "薪資 > " & minSalary, _
           vbInformation, "篩選完成"
End Sub

' 清除所有篩選條件並顯示全部資料
Sub ClearAllFilters()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    MsgBox "已清除所有篩選條件！", vbInformation, "完成"
End Sub

' 將篩選結果複製到新工作表
Sub CopyFilteredResultToSheet()
    Dim wsSource As Worksheet
    Dim wsResult As Worksheet
    Dim visibleRange As Range
    Dim lastRow As Long
    Dim lastCol As Long

    Set wsSource = ActiveSheet

    If Not wsSource.AutoFilterMode Then
        MsgBox "目前未套用篩選，請先執行篩選！", vbExclamation, "警告"
        Exit Sub
    End If

    Set wsResult = GetOrCreateSheet(ThisWorkbook, "篩選結果")

    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    Set visibleRange = wsSource.Range( _
        wsSource.Cells(1, 1), _
        wsSource.Cells(lastRow, lastCol)).SpecialCells(xlCellTypeVisible)

    visibleRange.Copy Destination:=wsResult.Cells(1, 1)
    wsResult.Columns.AutoFit
    wsResult.Activate
    MsgBox "篩選結果已複製到工作表「篩選結果」！", vbInformation, "完成"
End Sub

' 進階篩選：依條件範圍篩選並輸出到指定位置
Sub AdvancedFilterExample()
    Dim wsData As Worksheet
    Dim wsCriteria As Worksheet
    Dim wsOutput As Worksheet

    Set wsData = GetOrCreateSheet(ThisWorkbook, "進階篩選資料")
    Call FillEmployeeData(wsData)

    Set wsCriteria = GetOrCreateSheet(ThisWorkbook, "篩選條件")
    ' 條件：部門=業務部 OR 薪資>=60000（OR關係需在不同列）
    wsCriteria.Range("A1").Value = "部門"
    wsCriteria.Range("B1").Value = "薪資"
    wsCriteria.Range("A2").Value = "業務部"
    wsCriteria.Range("B3").Value = ">=60000"

    Set wsOutput = GetOrCreateSheet(ThisWorkbook, "進階篩選結果")

    wsData.Range("A1").CurrentRegion.AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=wsCriteria.Range("A1:B3"), _
        CopyToRange:=wsOutput.Range("A1"), _
        Unique:=False

    wsOutput.Columns.AutoFit
    wsOutput.Activate
    MsgBox "進階篩選完成！結果已輸出至「進階篩選結果」工作表。", vbInformation, "完成"
End Sub

' 填入員工資料
Private Sub FillEmployeeData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "薪資"
    ws.Range("D1").Value = "年資"
    ws.Range("A2").Value = "張三"
    ws.Range("B2").Value = "研發部"
    ws.Range("C2").Value = 70000
    ws.Range("D2").Value = 5
    ws.Range("A3").Value = "李四"
    ws.Range("B3").Value = "業務部"
    ws.Range("C3").Value = 55000
    ws.Range("D3").Value = 3
    ws.Range("A4").Value = "王五"
    ws.Range("B4").Value = "研發部"
    ws.Range("C4").Value = 80000
    ws.Range("D4").Value = 8
    ws.Range("A5").Value = "趙六"
    ws.Range("B5").Value = "行政部"
    ws.Range("C5").Value = 45000
    ws.Range("D5").Value = 2
    ws.Range("A6").Value = "孫七"
    ws.Range("B6").Value = "業務部"
    ws.Range("C6").Value = 62000
    ws.Range("D6").Value = 6
    ws.Range("A7").Value = "周八"
    ws.Range("B7").Value = "研發部"
    ws.Range("C7").Value = 50000
    ws.Range("D7").Value = 1
    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
