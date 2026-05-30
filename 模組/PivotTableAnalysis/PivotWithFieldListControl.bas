Attribute VB_Name = "PivotWithFieldListControl"
Option Explicit
'*************************************************************************************
'模組名稱: PivotWithFieldListControl
'功能說明: 以VBA動態控制樞紐分析表的欄位清單，新增列/欄/值欄位並套用樣式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestPivotWithFieldListControl()
    Call CreatePivotWithFieldControl
End Sub

' 建立具有欄位清單控制的樞紐分析表
Sub CreatePivotWithFieldControl()
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField

    On Error GoTo ErrHandler
    Set wb = ThisWorkbook

    Set wsData = GetOrCreateFieldSheet(wb, "部門銷售資料")
    Call FillDeptSalesData(wsData)

    Set wsPivot = GetOrCreateFieldSheet(wb, "欄位控制樞紐")

    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="部門樞紐")

    Set pf = pt.PivotFields("部門")
    pf.Orientation = xlRowField
    pf.Position = 1

    Set pf = pt.PivotFields("季度")
    pf.Orientation = xlColumnField
    pf.Position = 1

    Set pf = pt.PivotFields("銷售額")
    pf.Orientation = xlDataField
    pf.Function = xlSum
    pf.NumberFormat = "#,##0"
    pf.Name = "銷售額合計"

    Set pf = pt.PivotFields("件數")
    pf.Orientation = xlDataField
    pf.Function = xlSum
    pf.Name = "件數合計"

    pt.TableStyle2 = "PivotStyleMedium9"

    wsPivot.Range("A1").Value = "樞紐分析表欄位清單控制範例"
    wsPivot.Range("A1").Font.Bold = True
    wsPivot.Range("A1").Font.Size = 14
    wsPivot.Range("A2").Value = "列：部門  欄：季度  值：銷售額合計、件數合計"

    wsPivot.Columns.AutoFit
    wsPivot.Activate
    MsgBox "欄位清單控制樞紐分析表已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 填入部門銷售資料
Private Sub FillDeptSalesData(ByVal ws As Worksheet)
    Dim data As Variant
    Dim i As Integer

    ws.Cells.Clear
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "季度"
    ws.Range("C1").Value = "銷售額"
    ws.Range("D1").Value = "件數"
    ws.Range("A1:D1").Font.Bold = True

    data = Array( _
        Array("業務部", "Q1", 120000, 45), _
        Array("業務部", "Q2", 135000, 52), _
        Array("業務部", "Q3", 118000, 43), _
        Array("業務部", "Q4", 160000, 60), _
        Array("行銷部", "Q1", 80000, 30), _
        Array("行銷部", "Q2", 95000, 35), _
        Array("行銷部", "Q3", 88000, 32), _
        Array("行銷部", "Q4", 102000, 40), _
        Array("技術部", "Q1", 60000, 20), _
        Array("技術部", "Q2", 72000, 25), _
        Array("技術部", "Q3", 68000, 22), _
        Array("技術部", "Q4", 85000, 30))

    For i = 0 To UBound(data)
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
        ws.Cells(i + 2, 4).Value = data(i)(3)
    Next i
    ws.Columns.AutoFit
End Sub

' 取得或建立工作表並清除內容
Private Function GetOrCreateFieldSheet(ByVal wb As Workbook, _
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
    Set GetOrCreateFieldSheet = ws
End Function
