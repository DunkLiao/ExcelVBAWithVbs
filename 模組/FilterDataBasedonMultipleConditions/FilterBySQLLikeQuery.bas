Attribute VB_Name = "FilterBySQLLikeQuery"
Option Explicit
'*************************************************************************************
'模組名稱: FilterBySQLLikeQuery
'功能說明: 以SQL風格的LIKE語法（支援*萬用字元）篩選工作表資料並輸出結果
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestFilterBySQLLikeQuery()
    Call FilterWithLikeQuery
End Sub

' SQL LIKE 風格篩選
Sub FilterWithLikeQuery()
    Dim ws As Worksheet
    Dim wsResult As Worksheet
    Dim lngLastRow As Long
    Dim lngResultRow As Long
    Dim i As Long
    Dim sQuery As String
    Dim sColName As String
    Dim intColIndex As Integer
    Dim intCount As Integer
    Dim sCellVal As String

    On Error GoTo ErrHandler
    Set ws = GetOrCreateSQLSheet(ThisWorkbook, "銷售資料來源")
    Call FillSQLSampleData(ws)

    sQuery = InputBox( _
        "請輸入LIKE篩選條件（支援 * 萬用字元）：" & Chr(13) & _
        "例如：王* 表示以王開頭" & Chr(13) & _
        "      *部 表示以部結尾" & Chr(13) & _
        "      *業* 表示包含業字", _
        "SQL LIKE 篩選條件", "王*")

    If sQuery = "" Then
        MsgBox "已取消操作。", vbInformation
        Exit Sub
    End If

    sColName = InputBox( _
        "請輸入要篩選的欄位名稱：" & Chr(13) & _
        "可用欄位：姓名、部門、職位、城市", _
        "選擇篩選欄位", "姓名")

    If sColName = "" Then
        MsgBox "已取消操作。", vbInformation
        Exit Sub
    End If

    intColIndex = FindLikeColumnIndex(ws, sColName)
    If intColIndex = 0 Then
        MsgBox "找不到欄位「" & sColName & "」，請確認欄位名稱。", vbExclamation
        Exit Sub
    End If

    Set wsResult = GetOrCreateSQLSheet(ThisWorkbook, "篩選結果_LIKE")
    lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Rows(1).Copy Destination:=wsResult.Rows(1)
    wsResult.Rows(1).Font.Bold = True
    lngResultRow = 2
    intCount = 0

    Application.ScreenUpdating = False

    For i = 2 To lngLastRow
        sCellVal = CStr(ws.Cells(i, intColIndex).Value)
        If LikePatternMatch(sCellVal, sQuery) Then
            ws.Rows(i).Copy Destination:=wsResult.Rows(lngResultRow)
            lngResultRow = lngResultRow + 1
            intCount = intCount + 1
        End If
    Next i

    wsResult.Columns.AutoFit
    wsResult.Cells(lngResultRow + 1, 1).Value = _
        "篩選條件：" & sColName & " LIKE '" & sQuery & "'"
    wsResult.Cells(lngResultRow + 2, 1).Value = "共找到 " & intCount & " 筆資料"
    wsResult.Cells(lngResultRow + 1, 1).Font.Italic = True
    wsResult.Cells(lngResultRow + 2, 1).Font.Bold = True
    wsResult.Activate
    Application.ScreenUpdating = True
    MsgBox "篩選完成！共找到 " & intCount & " 筆符合條件的資料。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' VBA LIKE 萬用字元比對（* 代表任意字元）
Private Function LikePatternMatch(ByVal s As String, ByVal pattern As String) As Boolean
    LikePatternMatch = (s Like pattern)
End Function

' 找到欄位名稱的索引
Private Function FindLikeColumnIndex(ByVal ws As Worksheet, _
    ByVal colName As String) As Integer
    Dim lngLastCol As Long
    Dim j As Long
    lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For j = 1 To lngLastCol
        If ws.Cells(1, j).Value = colName Then
            FindLikeColumnIndex = CInt(j)
            Exit Function
        End If
    Next j
    FindLikeColumnIndex = 0
End Function

' 填入銷售範例資料
Private Sub FillSQLSampleData(ByVal ws As Worksheet)
    Dim data As Variant
    Dim i As Integer

    ws.Cells.Clear
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "職位"
    ws.Range("D1").Value = "城市"
    ws.Range("E1").Value = "薪資"
    ws.Range("A1:E1").Font.Bold = True

    data = Array( _
        Array("王小明", "業務部", "業務專員", "台北", 45000), _
        Array("陳美玲", "行銷部", "行銷企劃", "台中", 42000), _
        Array("林大偉", "技術部", "工程師", "台北", 55000), _
        Array("黃雅琪", "業務部", "業務主任", "高雄", 52000), _
        Array("李建宏", "人資部", "HR專員", "台北", 40000), _
        Array("張惠芳", "業務部", "業務副理", "台南", 58000), _
        Array("吳俊傑", "技術部", "資深工程師", "台北", 68000), _
        Array("劉怡君", "行銷部", "行銷主任", "台中", 50000), _
        Array("蔡宗翰", "業務部", "業務經理", "台北", 72000), _
        Array("許雅婷", "人資部", "HR主管", "高雄", 60000))

    For i = 0 To UBound(data)
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
        ws.Cells(i + 2, 4).Value = data(i)(3)
        ws.Cells(i + 2, 5).Value = data(i)(4)
        ws.Cells(i + 2, 5).NumberFormat = "#,##0"
    Next i
    ws.Columns.AutoFit
End Sub

' 取得或建立工作表並清除內容
Private Function GetOrCreateSQLSheet(ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSQLSheet = ws
End Function
