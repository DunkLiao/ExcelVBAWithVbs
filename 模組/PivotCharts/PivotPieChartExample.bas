Option Explicit

' 建立樞紐分析圓形圖，呈現各區銷售金額占比。
Public Sub CreatePivotPieChartExample()
    On Error GoTo ErrHandler

    Dim dataWs As Worksheet
    Dim pivotWs As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim chartObj As ChartObject

    Set dataWs = GetOrCreatePivotChartWorksheet("圓形樞紐資料")
    Set pivotWs = GetOrCreatePivotChartWorksheet("圓形樞紐圖")
    dataWs.Cells.Clear
    pivotWs.Cells.Clear

    Call FillPivotPieData(dataWs)
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataWs.Range("A1").CurrentRegion)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWs.Range("A3"), TableName:="PiePivotTable")

    With pivotTable
        .PivotFields("區域").Orientation = xlRowField
        .AddDataField .PivotFields("金額"), "金額合計", xlSum
        .DataFields(1).NumberFormat = "#,##0"
    End With

    Set chartObj = pivotWs.ChartObjects.Add(Left:=pivotWs.Range("D3").Left, Top:=pivotWs.Range("D3").Top, Width:=420, Height:=300)
    With chartObj.Chart
        .SetSourceData Source:=pivotTable.TableRange1
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "各區銷售占比"
        .SeriesCollection(1).ApplyDataLabels
    End With

    pivotWs.Columns.AutoFit
    MsgBox "樞紐分析圓形圖已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立樞紐分析圖失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillPivotPieData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("區域", "產品", "金額")
    ws.Range("A2:C2").Value = Array("北區", "A", 22000)
    ws.Range("A3:C3").Value = Array("北區", "B", 18000)
    ws.Range("A4:C4").Value = Array("中區", "A", 26000)
    ws.Range("A5:C5").Value = Array("中區", "B", 21000)
    ws.Range("A6:C6").Value = Array("南區", "A", 24000)
    ws.Range("A7:C7").Value = Array("南區", "B", 19500)
End Sub

Private Function GetOrCreatePivotChartWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreatePivotChartWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreatePivotChartWorksheet Is Nothing Then
        Set GetOrCreatePivotChartWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreatePivotChartWorksheet.Name = sheetName
    End If
End Function