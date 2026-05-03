Attribute VB_Name = "PercentagePivotTable"
Option Explicit
'*************************************************************************************
'模組名稱: PercentagePivotTable
'功能說明: 樞紐分析表值欄位顯示方式：佔總計百分比
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/3
'
'*************************************************************************************

' 測試入口
Sub TestPercentagePivotTable()
    Call CreatePercentagePivotTable
End Sub

' 建立百分比顯示的樞紐分析表
' 加入銷售額加總及佔總計百分比兩個值欄位
Sub CreatePercentagePivotTable()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim pfSum As PivotField
    Dim pfPct As PivotField

    ' 準備工作表
    Set wsData = GetOrCreateSheet(ThisWorkbook, "銷售資料")
    Call FillSalesData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "百分比樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="百分比樞紐分析表")

    ' 列欄位：地區
    With pt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 欄欄位：產品
    With pt.PivotFields("產品")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' 值欄位 1：銷售額加總（原始金額）
    Set pfSum = pt.AddDataField(pt.PivotFields("銷售額"), "銷售額加總", xlSum)
    pfSum.NumberFormat = "#,##0"

    ' 值欄位 2：銷售額佔總計百分比
    Set pfPct = pt.AddDataField(pt.PivotFields("銷售額"), "銷售額佔比", xlSum)
    pfPct.Calculation = xlPercentOfTotal
    pfPct.NumberFormat = "0.00%"

    wsPivot.Activate
    MsgBox "百分比樞紐分析表已建立完成！右側欄位顯示佔總計百分比。", vbInformation, "完成"
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
