Attribute VB_Name = "NumberFormatPivot"
Option Explicit
'*************************************************************************************
'模組名稱: NumberFormatPivot
'功能說明: 數值欄位格式設定示範
'          對樞紐分析表的值欄位套用
'          千分位、小數位數、百分比、貨幣符號等數字格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestNumberFormatPivot()
    Call CreateNumberFormatPivot
End Sub

' 建立多格式值欄位樞紐分析表
Sub CreateNumberFormatPivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim dfSales As PivotField
    Dim dfQty As PivotField
    Dim dfRate As PivotField

    Set wsData = GetOrCreateSheet(ThisWorkbook, "產品銷售")
    Call FillProductData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "數字格式樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="數字格式樞紐分析表")

    ' 列欄位：產品
    With pt.PivotFields("產品")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 值欄位 1：銷售額（貨幣格式）
    Set dfSales = pt.AddDataField(pt.PivotFields("銷售額"), "銷售額", xlSum)
    dfSales.NumberFormat = "NT$#,##0"

    ' 值欄位 2：數量（整數千分位）
    Set dfQty = pt.AddDataField(pt.PivotFields("數量"), "銷售數量", xlSum)
    dfQty.NumberFormat = "#,##0 件"

    ' 值欄位 3：退貨率（百分比）
    Set dfRate = pt.AddDataField(pt.PivotFields("退貨率"), "平均退貨率", xlAverage)
    dfRate.NumberFormat = "0.00%"

    ' 值欄位排列為欄方向
    pt.DataPivotField.Orientation = xlColumnField

    wsPivot.Columns("A:E").AutoFit
    wsPivot.Activate
    MsgBox "數字格式樞紐分析表已建立！" & Chr(13) & _
           "分別套用貨幣、千分位計數、百分比格式。", vbInformation, "完成"
End Sub

' 填入產品銷售資料（含退貨率）
Private Sub FillProductData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "銷售額"
    ws.Range("C1").Value = "數量"
    ws.Range("D1").Value = "退貨率"

    ws.Range("A2").Value = "筆記型電腦" : ws.Range("B2").Value = 120000 : ws.Range("C2").Value = 10 : ws.Range("D2").Value = 0.02
    ws.Range("A3").Value = "智慧型手機" : ws.Range("B3").Value = 85000  : ws.Range("C3").Value = 17 : ws.Range("D3").Value = 0.035
    ws.Range("A4").Value = "平板電腦"   : ws.Range("B4").Value = 63000  : ws.Range("C4").Value = 9  : ws.Range("D4").Value = 0.015
    ws.Range("A5").Value = "筆記型電腦" : ws.Range("B5").Value = 95000  : ws.Range("C5").Value = 8  : ws.Range("D5").Value = 0.025
    ws.Range("A6").Value = "智慧型手機" : ws.Range("B6").Value = 110000 : ws.Range("C6").Value = 22 : ws.Range("D6").Value = 0.04
    ws.Range("A7").Value = "耳機"       : ws.Range("B7").Value = 28000  : ws.Range("C7").Value = 35 : ws.Range("D7").Value = 0.05
    ws.Range("A8").Value = "耳機"       : ws.Range("B8").Value = 31000  : ws.Range("C8").Value = 40 : ws.Range("D8").Value = 0.045
    ws.Range("A9").Value = "平板電腦"   : ws.Range("B9").Value = 72000  : ws.Range("C9").Value = 12 : ws.Range("D9").Value = 0.02

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
