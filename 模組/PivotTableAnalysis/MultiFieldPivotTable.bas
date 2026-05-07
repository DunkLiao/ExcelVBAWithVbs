Attribute VB_Name = "MultiFieldPivotTable"
Option Explicit
'*************************************************************************************
'模組名稱: MultiFieldPivotTable
'功能說明: 多欄位列/欄分組樞紐分析表（列: 地區+產品，值: 銷售額+數量）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/3
'
'*************************************************************************************

' 測試入口
Sub TestMultiFieldPivotTable()
    Call CreateMultiFieldPivotTable
End Sub

' 建立多欄位樞紐分析表
' 列欄位: 地區、產品（兩層）/ 值: 銷售額加總、數量加總
Sub CreateMultiFieldPivotTable()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range

    ' 準備工作表
    Set wsData = GetOrCreateSheet(ThisWorkbook, "銷售資料")
    Call FillSalesData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "多欄位樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    ' 建立 PivotCache 與樞紐分析表
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="多欄位樞紐分析表")

    ' 第一層列欄位：地區
    With pt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 第二層列欄位：產品
    With pt.PivotFields("產品")
        .Orientation = xlRowField
        .Position = 2
    End With

    ' 值欄位 1：銷售額加總
    pt.AddDataField pt.PivotFields("銷售額"), "銷售額加總", xlSum

    ' 值欄位 2：數量加總
    pt.AddDataField pt.PivotFields("數量"), "數量加總", xlSum

    ' 設定多欄位佈局（壓縮模式）
    pt.RowAxisLayout xlCompactRow

    wsPivot.Activate
    MsgBox "多欄位樞紐分析表已建立完成！", vbInformation, "完成"
End Sub

' 填入範例銷售資料（共 8 筆：日期、地區、產品、數量、單價、銷售額）
Private Sub FillSalesData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "地區"
    ws.Range("C1").Value = "產品"
    ws.Range("D1").Value = "數量"
    ws.Range("E1").Value = "單價"
    ws.Range("F1").Value = "銷售額"

    ws.Range("A2").Value = DateSerial(2026, 1, 5)
    ws.Range("B2").Value = "北部"
    ws.Range("C2").Value = "產品A"
    ws.Range("D2").Value = 100
    ws.Range("E2").Value = 50
    ws.Range("F2").Value = 5000

    ws.Range("A3").Value = DateSerial(2026, 1, 10)
    ws.Range("B3").Value = "南部"
    ws.Range("C3").Value = "產品B"
    ws.Range("D3").Value = 80
    ws.Range("E3").Value = 75
    ws.Range("F3").Value = 6000

    ws.Range("A4").Value = DateSerial(2026, 2, 5)
    ws.Range("B4").Value = "北部"
    ws.Range("C4").Value = "產品B"
    ws.Range("D4").Value = 60
    ws.Range("E4").Value = 75
    ws.Range("F4").Value = 4500

    ws.Range("A5").Value = DateSerial(2026, 2, 15)
    ws.Range("B5").Value = "中部"
    ws.Range("C5").Value = "產品A"
    ws.Range("D5").Value = 120
    ws.Range("E5").Value = 50
    ws.Range("F5").Value = 6000

    ws.Range("A6").Value = DateSerial(2026, 3, 1)
    ws.Range("B6").Value = "南部"
    ws.Range("C6").Value = "產品A"
    ws.Range("D6").Value = 90
    ws.Range("E6").Value = 50
    ws.Range("F6").Value = 4500

    ws.Range("A7").Value = DateSerial(2026, 3, 20)
    ws.Range("B7").Value = "中部"
    ws.Range("C7").Value = "產品B"
    ws.Range("D7").Value = 150
    ws.Range("E7").Value = 75
    ws.Range("F7").Value = 11250

    ws.Range("A8").Value = DateSerial(2026, 4, 5)
    ws.Range("B8").Value = "北部"
    ws.Range("C8").Value = "產品A"
    ws.Range("D8").Value = 200
    ws.Range("E8").Value = 50
    ws.Range("F8").Value = 10000

    ws.Range("A9").Value = DateSerial(2026, 4, 18)
    ws.Range("B9").Value = "南部"
    ws.Range("C9").Value = "產品B"
    ws.Range("D9").Value = 110
    ws.Range("E9").Value = 75
    ws.Range("F9").Value = 8250

    ws.Columns("A").NumberFormat = "yyyy/m/d"
    ws.Columns("A:F").AutoFit
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
