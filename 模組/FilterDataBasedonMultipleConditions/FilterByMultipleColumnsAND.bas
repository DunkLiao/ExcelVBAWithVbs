Attribute VB_Name = "FilterByMultipleColumnsAND"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByMultipleColumnsAND
'功能說明: 依多個欄位的 AND 條件同時篩選資料，將符合結果複製至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestFilterByMultipleColumnsAND()
    Call CreateFilterByMultipleColumnsANDExample
End Sub

' 建立範例資料並執行多欄 AND 篩選
Sub CreateFilterByMultipleColumnsANDExample()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim names As Variant
    Dim depts As Variant
    Dim salaries As Variant
    Dim years As Variant
    Dim i As Integer

    Set ws = GetOrCreateSheet(ThisWorkbook, "多欄AND篩選資料")

    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "薪資"
    ws.Range("D1").Value = "年資"
    ws.Range("A1:D1").Font.Bold = True

    names = Array("張明", "李華", "王芳", "陳建", "林靜", "黃強", "吳敏", "趙偉")
    depts = Array("業務", "工程", "業務", "行銷", "工程", "業務", "行銷", "工程")
    salaries = Array(45000, 72000, 38000, 55000, 80000, 50000, 42000, 68000)
    years = Array(3, 7, 1, 5, 9, 4, 2, 6)

    For i = 0 To UBound(names)
        ws.Cells(i + 2, 1).Value = names(i)
        ws.Cells(i + 2, 2).Value = depts(i)
        ws.Cells(i + 2, 3).Value = salaries(i)
        ws.Cells(i + 2, 4).Value = years(i)
    Next i

    ws.Columns("A:D").AutoFit

    Call FilterByMultipleColumnsAND(ws)
    Exit Sub

ErrorHandler:
    MsgBox "建立範例時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 依多欄 AND 條件篩選資料
Sub FilterByMultipleColumnsAND(ByVal wsSource As Worksheet)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim wsDest As Worksheet
    Dim destRow As Long
    Dim i As Long
    Dim dept As String
    Dim minSalary As Double
    Dim minYears As Integer
    Dim rowDept As String
    Dim rowSalary As Double
    Dim rowYears As Integer
    Dim resultCount As Long

    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    ' 設定篩選條件（部門=業務 AND 薪資>=45000 AND 年資>=3）
    dept = "業務"
    minSalary = 45000
    minYears = 3

    Set wsDest = GetOrCreateSheet(ThisWorkbook, "AND篩選結果")

    wsSource.Rows(1).Copy Destination:=wsDest.Rows(1)
    wsDest.Rows(1).Font.Bold = True
    destRow = 2

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        rowDept = CStr(wsSource.Cells(i, 2).Value)
        If IsNumeric(wsSource.Cells(i, 3).Value) Then
            rowSalary = CDbl(wsSource.Cells(i, 3).Value)
        Else
            rowSalary = 0
        End If
        If IsNumeric(wsSource.Cells(i, 4).Value) Then
            rowYears = CInt(wsSource.Cells(i, 4).Value)
        Else
            rowYears = 0
        End If

        ' AND 多重條件判斷
        If rowDept = dept And rowSalary >= minSalary And rowYears >= minYears Then
            wsSource.Rows(i).Copy Destination:=wsDest.Rows(destRow)
            destRow = destRow + 1
        End If
    Next i

    wsDest.Columns.AutoFit
    Application.ScreenUpdating = True

    resultCount = destRow - 2
    MsgBox "AND 多欄篩選完成！" & vbCrLf & _
           "條件：部門=" & dept & "，薪資>=" & minSalary & "，年資>=" & minYears & vbCrLf & _
           "符合筆數：" & resultCount & " 筆", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "篩選時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
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
