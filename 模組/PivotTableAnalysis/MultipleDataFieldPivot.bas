Option Explicit
Attribute VB_Name = "MultipleDataFieldPivot"
'*************************************************************************************
'模組名稱: MultipleDataFieldPivot
'功能說明: 建立含多個資料欄位的樞紐分析表，同時顯示加總、平均等統計值
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestMultipleDataFieldPivot()
    Call CreateMultiDataFieldPivot(ThisWorkbook)
End Sub

Sub CreateMultiDataFieldPivot(ByVal wb As Workbook)
    On Error GoTo ErrorHandler

    Dim dataWs As Worksheet
    Set dataWs = GetOrCreateMDFSheet(wb, "銷售原始資料")
    dataWs.Cells.Clear
    Call FillMultiPivotData(dataWs)

    Dim pivotWs As Worksheet
    Set pivotWs = GetOrCreateMDFSheet(wb, "多欄位樞紐")
    pivotWs.Cells.Clear

    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataWs.Range("A1").CurrentRegion)

    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="多欄位樞紐表")

    pt.PivotFields("產品類別").Orientation = xlRowField
    pt.PivotFields("產品類別").Position = 1
    pt.PivotFields("銷售月份").Orientation = xlColumnField
    pt.PivotFields("銷售月份").Position = 1

    Dim pfAmt As PivotField
    Set pfAmt = pt.PivotFields("銷售金額")
    pfAmt.Orientation = xlDataField
    pfAmt.Function = xlSum
    pfAmt.Name = "加總:銷售金額"
    pfAmt.NumberFormat = "#,##0"

    Dim pfQty As PivotField
    Set pfQty = pt.PivotFields("銷售數量")
    pfQty.Orientation = xlDataField
    pfQty.Function = xlAverage
    pfQty.Name = "平均:銷售數量"
    pfQty.NumberFormat = "0.00"

    pivotWs.Columns.AutoFit
    MsgBox "多欄位樞紐分析表已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立樞紐分析表時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillMultiPivotData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("產品類別", "銷售月份", "銷售金額", "銷售數量")

    Dim categories As Variant
    categories = Array("電子產品", "生活用品", "電子產品", "服飾", _
        "生活用品", "服飾", "電子產品", "生活用品")

    Dim months As Variant
    months = Array("一月", "一月", "二月", "一月", _
        "二月", "二月", "三月", "三月")

    Dim amounts As Variant
    amounts = Array(50000, 12000, 65000, 18000, _
        15000, 22000, 72000, 9500)

    Dim quantities As Variant
    quantities = Array(10, 30, 13, 45, 38, 55, 14, 25)

    Dim r As Integer
    For r = 1 To 8
        ws.Cells(r + 1, 1).Value = categories(r - 1)
        ws.Cells(r + 1, 2).Value = months(r - 1)
        ws.Cells(r + 1, 3).Value = amounts(r - 1)
        ws.Cells(r + 1, 4).Value = quantities(r - 1)
    Next r

    ws.Columns.AutoFit
End Sub

Private Function GetOrCreateMDFSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateMDFSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateMDFSheet Is Nothing Then
        Set GetOrCreateMDFSheet = wb.Worksheets.Add
        GetOrCreateMDFSheet.Name = sheetName
    End If
End Function
