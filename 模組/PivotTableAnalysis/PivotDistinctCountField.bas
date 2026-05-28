Attribute VB_Name = "PivotDistinctCountField"
Option Explicit
'*************************************************************************************
'模組名稱: PivotDistinctCountField
'功能說明: 建立包含計數與加總欄位的樞紐分析表，示範不重複計數的應用方式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestPivotDistinctCount()
    Call CreateDistinctCountDemo(ThisWorkbook)
End Sub

Sub CreateDistinctCountDemo(ByVal wb As Workbook)
    Dim dataWs    As Worksheet
    Dim pivotWs   As Worksheet
    Dim pc        As PivotCache
    Dim pt        As PivotTable

    On Error Resume Next
    Set dataWs = wb.Worksheets("銷售明細")
    On Error GoTo 0
    If dataWs Is Nothing Then
        Set dataWs = wb.Worksheets.Add
        dataWs.Name = "銷售明細"
    End If
    dataWs.Cells.Clear
    Call FillDistinctCountData(dataWs)

    On Error Resume Next
    Set pivotWs = wb.Worksheets("不重複計數樞紐")
    On Error GoTo 0
    If pivotWs Is Nothing Then
        Set pivotWs = wb.Worksheets.Add(After:=dataWs)
        pivotWs.Name = "不重複計數樞紐"
    End If
    pivotWs.Cells.Clear

    On Error GoTo ErrHandler
    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataWs.UsedRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="不重複計數樞紐")

    With pt.PivotFields("部門")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields("訂單金額")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0"
        .Name = "訂單金額合計"
    End With

    With pt.PivotFields("客戶名稱")
        .Orientation = xlDataField
        .Function = xlCount
        .Name = "訂單筆數"
    End With

    pivotWs.Range("A1").Value = "各部門銷售摘要（含訂單筆數）"
    pivotWs.Range("A1").Font.Bold = True
    pivotWs.Range("A1").Font.Size = 14
    pivotWs.Range("A2").Value = "＊如需不重複客戶數，請啟用資料模型並改用 DistinctCount 彙總函數"
    pivotWs.Range("A2").Font.Color = RGB(150, 0, 0)
    pivotWs.Columns.AutoFit
    pivotWs.Activate
    MsgBox "不重複計數樞紐分析表已建立完畢！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立樞紐失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillDistinctCountData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("部門", "客戶名稱", "產品", "訂單金額")
    ws.Range("A2:D2").Value = Array("業務部", "客戶A", "產品甲", 15000)
    ws.Range("A3:D3").Value = Array("業務部", "客戶B", "產品乙", 23000)
    ws.Range("A4:D4").Value = Array("業務部", "客戶A", "產品丙", 8000)
    ws.Range("A5:D5").Value = Array("工程部", "客戶C", "產品甲", 42000)
    ws.Range("A6:D6").Value = Array("工程部", "客戶D", "產品丁", 31000)
    ws.Range("A7:D7").Value = Array("工程部", "客戶C", "產品乙", 18000)
    ws.Range("A8:D8").Value = Array("財務部", "客戶E", "產品丙", 9500)
    ws.Range("A9:D9").Value = Array("業務部", "客戶B", "產品丁", 27000)
    ws.Range("A10:D10").Value = Array("財務部", "客戶F", "產品甲", 12000)
    ws.Range("A11:D11").Value = Array("工程部", "客戶D", "產品丙", 35000)
    ws.Columns("A:D").AutoFit
End Sub
