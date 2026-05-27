Option Explicit
Attribute VB_Name = "FilterByCheckboxList"
'*************************************************************************************
'模組名稱: 依核取方塊清單篩選資料
'功能說明: 讀取工作表「篩選設定」B欄中值為 TRUE 的項目，
'          以 AutoFilter 篩選主資料表，結果複製到新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub TestFilterByCheckboxList()
    Call SetupCheckboxListDemo
    Call FilterByCheckboxList("資料表", "篩選設定", 1)
End Sub

Sub SetupCheckboxListDemo()
    Dim wsData As Worksheet
    Dim wsFilter As Worksheet

    Set wsData = GetOrCreateWsFilter("資料表")
    wsData.Cells.Clear
    wsData.Range("A1").Value = "部門"
    wsData.Range("B1").Value = "姓名"
    wsData.Range("C1").Value = "業績"

    Dim depts As Variant
    Dim names As Variant
    Dim scores As Variant
    depts  = Array("業務部", "財務部", "業務部", "研發部", "業務部", "財務部")
    names  = Array("王大明", "李小華", "張美玲", "陳建宏", "林淑芬", "黃志偉")
    scores = Array(320000, 150000, 480000, 210000, 390000, 175000)

    Dim i As Integer
    For i = 0 To 5
        wsData.Cells(i + 2, 1).Value = depts(i)
        wsData.Cells(i + 2, 2).Value = names(i)
        wsData.Cells(i + 2, 3).Value = scores(i)
    Next i
    wsData.Columns("A:C").AutoFit

    Set wsFilter = GetOrCreateWsFilter("篩選設定")
    wsFilter.Cells.Clear
    wsFilter.Range("A1").Value = "部門"
    wsFilter.Range("B1").Value = "篩選"
    wsFilter.Range("A2").Value = "業務部"
    wsFilter.Range("B2").Value = True
    wsFilter.Range("A3").Value = "財務部"
    wsFilter.Range("B3").Value = False
    wsFilter.Range("A4").Value = "研發部"
    wsFilter.Range("B4").Value = True
    wsFilter.Columns("A:B").AutoFit
End Sub

Sub FilterByCheckboxList( _
    ByVal dataSheetName As String, _
    ByVal filterSheetName As String, _
    ByVal filterColIndex As Long)

    On Error GoTo ErrorHandler

    Dim wsData As Worksheet
    Dim wsFilter As Worksheet
    Dim wsResult As Worksheet
    Dim lastFilterRow As Long
    Dim i As Long
    Dim itemCount As Integer
    Dim itemArr() As String

    Set wsData = ThisWorkbook.Worksheets(dataSheetName)
    Set wsFilter = ThisWorkbook.Worksheets(filterSheetName)

    ' 收集勾選項目
    lastFilterRow = wsFilter.Cells(wsFilter.Rows.Count, 1).End(xlUp).Row
    itemCount = 0
    ReDim itemArr(0)

    For i = 2 To lastFilterRow
        If wsFilter.Cells(i, 2).Value = True Then
            ReDim Preserve itemArr(itemCount)
            itemArr(itemCount) = CStr(wsFilter.Cells(i, 1).Value)
            itemCount = itemCount + 1
        End If
    Next i

    If itemCount = 0 Then
        MsgBox "沒有任何勾選項目，請在篩選設定工作表勾選至少一項。", vbExclamation, "提示"
        Exit Sub
    End If

    ' 套用 AutoFilter
    wsData.AutoFilterMode = False
    wsData.Range("A1").AutoFilter Field:=filterColIndex, _
        Criteria1:=itemArr, Operator:=xlFilterValues

    ' 複製到結果工作表
    Set wsResult = GetOrCreateWsFilter("篩選結果")
    wsResult.Cells.Clear

    Dim visibleRange As Range
    On Error Resume Next
    Set visibleRange = wsData.UsedRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo ErrorHandler

    If Not visibleRange Is Nothing Then
        visibleRange.Copy Destination:=wsResult.Range("A1")
    End If

    wsData.AutoFilterMode = False
    wsResult.Columns.AutoFit

    MsgBox "篩選完成！已複製結果到「篩選結果」工作表。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wsData Is Nothing Then wsData.AutoFilterMode = False
    MsgBox "篩選時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetOrCreateWsFilter(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWsFilter = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWsFilter Is Nothing Then
        Set GetOrCreateWsFilter = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateWsFilter.Name = sheetName
    End If
End Function
