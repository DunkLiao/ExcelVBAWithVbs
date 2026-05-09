Attribute VB_Name = "SplitSheetToNewSheets"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetToNewSheets
'功能說明: 依指定欄位的唯一值，將工作表資料拆分為同一活頁簿中的多個工作表
'          （不另存新檔，直接在目前活頁簿新增分頁）
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例資料後依部門拆分為多個工作表
Sub TestSplitToSheets()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheetInBook(ThisWorkbook, "待分割來源")
    Call FillSampleSheetData(ws)
    Call SplitSheetIntoNewSheets(ws, 2)
End Sub

' 依欄位唯一值將資料拆分至同活頁簿的新工作表
' ws: 來源工作表  splitColIndex: 分割依據欄號
Sub SplitSheetIntoNewSheets(ByVal ws As Worksheet, ByVal splitColIndex As Integer)
    Dim lastRow As Long
    Dim i As Long
    Dim colCount As Long
    Dim cellValue As String
    Dim uniqueValues As Collection
    Dim alreadyAdded As Boolean
    Dim key As Variant
    Dim wsNew As Worksheet
    Dim newRow As Long
    Dim copyRow As Long

    lastRow = ws.Cells(ws.Rows.Count, splitColIndex).End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "工作表無資料可分割！", vbExclamation, "警告"
        Exit Sub
    End If

    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set uniqueValues = New Collection
    For i = 2 To lastRow
        cellValue = CStr(ws.Cells(i, splitColIndex).Value)
        alreadyAdded = False
        For Each key In uniqueValues
            If key = cellValue Then
                alreadyAdded = True
                Exit For
            End If
        Next key
        If Not alreadyAdded Then uniqueValues.Add cellValue
    Next i

    Application.ScreenUpdating = False

    For Each key In uniqueValues
        On Error Resume Next
        Application.DisplayAlerts = False
        ws.Parent.Worksheets(CStr(key)).Delete
        Application.DisplayAlerts = True
        On Error GoTo 0

        Set wsNew = ws.Parent.Worksheets.Add( _
            After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
        wsNew.Name = Left(CStr(key), 31)
        ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Copy _
            Destination:=wsNew.Cells(1, 1)
        newRow = 2
        For copyRow = 2 To lastRow
            If CStr(ws.Cells(copyRow, splitColIndex).Value) = CStr(key) Then
                ws.Range(ws.Cells(copyRow, 1), ws.Cells(copyRow, colCount)).Copy _
                    Destination:=wsNew.Cells(newRow, 1)
                newRow = newRow + 1
            End If
        Next copyRow
        wsNew.Columns.AutoFit
    Next key

    Application.ScreenUpdating = True
    MsgBox "已在目前活頁簿建立 " & uniqueValues.Count & " 個新工作表！", vbInformation, "完成"
End Sub

' 建立範例資料
Private Sub FillSampleSheetData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "職稱"
    ws.Range("A2").Value = "王大明": ws.Range("B2").Value = "業務部": ws.Range("C2").Value = "業務專員"
    ws.Range("A3").Value = "李小美": ws.Range("B3").Value = "研發部": ws.Range("C3").Value = "工程師"
    ws.Range("A4").Value = "張志偉": ws.Range("B4").Value = "業務部": ws.Range("C4").Value = "業務主任"
    ws.Range("A5").Value = "陳美如": ws.Range("B5").Value = "行政部": ws.Range("C5").Value = "行政助理"
    ws.Range("A6").Value = "林正雄": ws.Range("B6").Value = "研發部": ws.Range("C6").Value = "研發主任"
    ws.Range("A7").Value = "吳淑芬": ws.Range("B7").Value = "行政部": ws.Range("C7").Value = "人事專員"
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetInBook(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetInBook = ws
End Function
