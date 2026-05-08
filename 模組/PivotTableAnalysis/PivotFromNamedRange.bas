Option Explicit

' 使用命名範圍建立樞紐分析表，示範來源範圍命名與彙總欄位設定。
Public Sub CreatePivotFromNamedRangeExample()
    On Error GoTo ErrHandler

    Dim dataWs As Worksheet
    Dim pivotWs As Worksheet
    Dim sourceRange As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable

    Set dataWs = GetOrCreatePivotWorksheet("命名範圍資料")
    Set pivotWs = GetOrCreatePivotWorksheet("命名範圍樞紐")
    dataWs.Cells.Clear
    pivotWs.Cells.Clear

    Call FillPivotNamedRangeData(dataWs)
    Set sourceRange = dataWs.Range("A1").CurrentRegion
    ThisWorkbook.Names.Add Name:="SalesNamedRange", RefersTo:=sourceRange

    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="SalesNamedRange")
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWs.Range("A3"), TableName:="NamedRangePivot")

    With pivotTable
        .PivotFields("區域").Orientation = xlRowField
        .PivotFields("產品").Orientation = xlColumnField
        .AddDataField .PivotFields("金額"), "金額合計", xlSum
        .DataFields(1).NumberFormat = "#,##0"
        .RowAxisLayout xlTabularRow
    End With

    pivotWs.Columns.AutoFit
    MsgBox "命名範圍樞紐分析表已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立樞紐分析表失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillPivotNamedRangeData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("區域", "產品", "業務", "金額")
    ws.Range("A2:D2").Value = Array("北區", "A", "王小明", 12000)
    ws.Range("A3:D3").Value = Array("北區", "B", "王小明", 9500)
    ws.Range("A4:D4").Value = Array("中區", "A", "陳美華", 13800)
    ws.Range("A5:D5").Value = Array("中區", "B", "陳美華", 11200)
    ws.Range("A6:D6").Value = Array("南區", "A", "林志強", 15600)
    ws.Range("A7:D7").Value = Array("南區", "B", "林志強", 10100)
End Sub

Private Function GetOrCreatePivotWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreatePivotWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreatePivotWorksheet Is Nothing Then
        Set GetOrCreatePivotWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreatePivotWorksheet.Name = sheetName
    End If
End Function