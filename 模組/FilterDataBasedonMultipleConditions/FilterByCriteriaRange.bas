Attribute VB_Name = "FilterByCriteriaRange"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByCriteriaRange
'功能說明: 使用進階篩選的條件範圍，依多欄組合條件篩選資料並輸出至指定位置
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestFilterByCriteriaRange()
    Call FilterByCriteriaRangeExample
End Sub

' 使用條件範圍進行進階篩選
Sub FilterByCriteriaRangeExample()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim dataWs As Worksheet
    Dim outputWs As Worksheet

    Set wb = ThisWorkbook
    Set dataWs = GetOrCreateSheet(wb, "篩選來源資料")
    Set outputWs = GetOrCreateSheet(wb, "條件範圍篩選結果")

    Call FillFilterSourceData(dataWs)
    Call SetupCriteriaAndFilter(dataWs, outputWs)

    outputWs.Activate
    MsgBox "條件範圍進階篩選完成！" & vbCrLf & _
           "篩選條件：部門=業務部 且 金額>=15000，" & vbCrLf & _
           "或 部門=研發部 且 金額>=20000", _
           vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "篩選時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 設定條件範圍並執行進階篩選
Private Sub SetupCriteriaAndFilter(ByVal dataWs As Worksheet, ByVal outputWs As Worksheet)
    Dim lastRow As Long
    Dim dataRange As Range
    Dim criteriaRange As Range
    Dim outputRange As Range

    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row
    Set dataRange = dataWs.Range("A1:D" & lastRow)

    Dim criteriaStartCell As Range
    Set criteriaStartCell = dataWs.Range("F1")

    criteriaStartCell.Value = "部門"
    criteriaStartCell.Offset(0, 1).Value = "銷售金額"

    criteriaStartCell.Offset(1, 0).Value = "業務部"
    criteriaStartCell.Offset(1, 1).Value = ">=15000"

    criteriaStartCell.Offset(2, 0).Value = "研發部"
    criteriaStartCell.Offset(2, 1).Value = ">=20000"

    Set criteriaRange = dataWs.Range("F1:G3")
    Set outputRange = outputWs.Range("A1")

    dataRange.AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=criteriaRange, _
        CopyToRange:=outputRange, _
        Unique:=False

    outputWs.Columns.AutoFit

    Dim resultLastRow As Long
    resultLastRow = outputWs.Cells(outputWs.Rows.Count, 1).End(xlUp).Row
    outputWs.Range("A" & (resultLastRow + 2)).Value = _
        "篩選條件：（部門=業務部 且 金額>=15000）或（部門=研發部 且 金額>=20000）"
    outputWs.Range("A" & (resultLastRow + 2)).Font.Italic = True
    outputWs.Range("A" & (resultLastRow + 2)).Font.Color = RGB(128, 128, 128)
End Sub

' 填入篩選來源範例資料
Private Sub FillFilterSourceData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "員工"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "月份"
    ws.Range("D1").Value = "銷售金額"
    ws.Range("A1:D1").Font.Bold = True

    Dim data As Variant
    data = Array( _
        Array("王小明", "業務部", "1月", 18000), _
        Array("李大華", "研發部", "1月", 22000), _
        Array("陳美玲", "業務部", "1月", 12000), _
        Array("張志偉", "行政部", "1月", 8500), _
        Array("林雅婷", "業務部", "2月", 25000), _
        Array("吳建國", "研發部", "2月", 17000), _
        Array("黃淑芬", "業務部", "2月", 9800), _
        Array("楊明哲", "行政部", "2月", 7600), _
        Array("蔡雅文", "研發部", "3月", 31000), _
        Array("許俊傑", "業務部", "3月", 16500) _
    )

    Dim i As Integer
    For i = 0 To 9
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
        ws.Cells(i + 2, 4).Value = data(i)(3)
    Next i

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
