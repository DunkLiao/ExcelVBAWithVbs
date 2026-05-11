Attribute VB_Name = "MultipleSummaryFunctionPivot"
Option Explicit
'*************************************************************************************
'模組名稱: MultipleSummaryFunctionPivot
'功能說明: 在樞紐分析表中為同一欄位設定多個摘要函數（加總、平均、最大、最小）的範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestMultipleSummaryFunctionPivot()
    Call CreateMultipleSummaryPivot
End Sub

' 建立含多重摘要函數的樞紐分析表
Sub CreateMultipleSummaryPivot()
    On Error GoTo ErrorHandler

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pvf As PivotField
    Dim dataRange As Range
    Dim lastRow As Long
    Dim idx As Integer

    ' 準備來源資料
    Set wsData = GetOrCreateWorksheetMSP("多摘要來源")
    wsData.Cells.Clear
    Call FillMultipleSummaryData(wsData)

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set dataRange = wsData.Range("A1").Resize(lastRow, 4)

    ' 準備樞紐輸出工作表
    Set wsPivot = GetOrCreateWorksheetMSP("多摘要函數樞紐")
    wsPivot.Cells.Clear

    ' 建立樞紐快取
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' 建立樞紐分析表
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="多摘要樞紐")

    ' 設定列欄位
    With pt.PivotFields("部門")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 加入四個值欄位，使用不同摘要函數
    Dim dataFieldNames(3) As String
    Dim dataFieldFunctions(3) As Long
    Dim dataFieldLabels(3) As String

    dataFieldNames(0) = "銷售額" : dataFieldFunctions(0) = xlSum : dataFieldLabels(0) = "加總-銷售額"
    dataFieldNames(1) = "銷售額" : dataFieldFunctions(1) = xlAverage : dataFieldLabels(1) = "平均-銷售額"
    dataFieldNames(2) = "銷售額" : dataFieldFunctions(2) = xlMax : dataFieldLabels(2) = "最大-銷售額"
    dataFieldNames(3) = "銷售額" : dataFieldFunctions(3) = xlMin : dataFieldLabels(3) = "最小-銷售額"

    For idx = 0 To 3
        Set pvf = pt.AddDataField( _
            pt.PivotFields(dataFieldNames(idx)), _
            dataFieldLabels(idx), _
            dataFieldFunctions(idx))
        pvf.NumberFormat = "#,##0"
    Next idx

    pt.HasAutoFormat = False
    wsPivot.Columns("A:F").AutoFit
    wsPivot.Activate

    MsgBox "多重摘要函數樞紐分析表已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立樞紐分析表時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetOrCreateWorksheetMSP(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheetMSP = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateWorksheetMSP Is Nothing Then
        Set GetOrCreateWorksheetMSP = ThisWorkbook.Worksheets.Add
        GetOrCreateWorksheetMSP.Name = sheetName
    End If
End Function

Private Sub FillMultipleSummaryData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "員工"
    ws.Range("C1").Value = "月份"
    ws.Range("D1").Value = "銷售額"

    Dim r As Integer
    Dim data(1 To 12, 1 To 4) As Variant
    data(1, 1) = "業務部" : data(1, 2) = "王大明" : data(1, 3) = "1月" : data(1, 4) = 120000
    data(2, 1) = "業務部" : data(2, 2) = "陳小華" : data(2, 3) = "1月" : data(2, 4) = 98000
    data(3, 1) = "行銷部" : data(3, 2) = "林美麗" : data(3, 3) = "1月" : data(3, 4) = 55000
    data(4, 1) = "行銷部" : data(4, 2) = "張志明" : data(4, 3) = "1月" : data(4, 4) = 62000
    data(5, 1) = "業務部" : data(5, 2) = "王大明" : data(5, 3) = "2月" : data(5, 4) = 135000
    data(6, 1) = "業務部" : data(6, 2) = "陳小華" : data(6, 3) = "2月" : data(6, 4) = 110000
    data(7, 1) = "行銷部" : data(7, 2) = "林美麗" : data(7, 3) = "2月" : data(7, 4) = 60000
    data(8, 1) = "行銷部" : data(8, 2) = "張志明" : data(8, 3) = "2月" : data(8, 4) = 72000
    data(9, 1) = "業務部" : data(9, 2) = "王大明" : data(9, 3) = "3月" : data(9, 4) = 158000
    data(10, 1) = "業務部" : data(10, 2) = "陳小華" : data(10, 3) = "3月" : data(10, 4) = 125000
    data(11, 1) = "行銷部" : data(11, 2) = "林美麗" : data(11, 3) = "3月" : data(11, 4) = 68000
    data(12, 1) = "行銷部" : data(12, 2) = "張志明" : data(12, 3) = "3月" : data(12, 4) = 80000

    For r = 1 To 12
        ws.Cells(r + 1, 1).Value = data(r, 1)
        ws.Cells(r + 1, 2).Value = data(r, 2)
        ws.Cells(r + 1, 3).Value = data(r, 3)
        ws.Cells(r + 1, 4).Value = data(r, 4)
    Next r

    ws.Columns("A:D").AutoFit
End Sub
