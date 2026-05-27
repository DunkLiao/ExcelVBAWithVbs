Option Explicit
Attribute VB_Name = "PivotWithCalculatedItem"
'*************************************************************************************
'模組名稱: 樞紐分析表計算項目範例
'功能說明: 建立包含計算項目（Calculated Item）的樞紐分析表，
'          示範在列欄位中新增自訂運算項目
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub TestPivotWithCalculatedItem()
    Call CreatePivotWithCalculatedItem("計算項目範例")
End Sub

Sub CreatePivotWithCalculatedItem(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache
    Dim dataRange As Range

    ' 建立資料工作表
    Set wsData = GetOrCreateWsCalcItem(sheetName & "_資料")
    wsData.Cells.Clear
    Call FillCalcItemData(wsData)

    ' 建立樞紐工作表
    Set wsPivot = GetOrCreateWsCalcItem(sheetName & "_樞紐")
    wsPivot.Cells.Clear

    Set dataRange = wsData.Range("A1:C13")
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="CalcItemPivot")

    With pt
        .PivotFields("區域").Orientation = xlRowField
        .PivotFields("區域").Position = 1

        .PivotFields("季度").Orientation = xlColumnField
        .PivotFields("季度").Position = 1

        .AddDataField .PivotFields("銷售額"), "加總-銷售額", xlSum

        ' 新增計算項目：上半年 = Q1 + Q2
        Dim pf As PivotField
        Set pf = pt.PivotFields("季度")
        pf.CalculatedItems.Add "上半年", "='Q1'+'Q2'"
    End With

    wsPivot.Columns.AutoFit
    MsgBox "含計算項目的樞紐分析表已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立樞紐分析表時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetOrCreateWsCalcItem(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWsCalcItem = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWsCalcItem Is Nothing Then
        Set GetOrCreateWsCalcItem = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateWsCalcItem.Name = sheetName
    End If
End Function

Private Sub FillCalcItemData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "區域"
    ws.Range("B1").Value = "季度"
    ws.Range("C1").Value = "銷售額"

    Dim areas As Variant
    Dim quarters As Variant
    Dim amounts As Variant
    areas = Array("北區", "北區", "北區", "北區", "南區", "南區", "南區", "南區", "中區", "中區", "中區", "中區")
    quarters = Array("Q1", "Q2", "Q3", "Q4", "Q1", "Q2", "Q3", "Q4", "Q1", "Q2", "Q3", "Q4")
    amounts = Array(120, 135, 150, 170, 100, 115, 130, 145, 90, 108, 122, 138)

    Dim i As Integer
    For i = 0 To 11
        ws.Cells(i + 2, 1).Value = areas(i)
        ws.Cells(i + 2, 2).Value = quarters(i)
        ws.Cells(i + 2, 3).Value = amounts(i)
    Next i

    ws.Columns("A:C").AutoFit
End Sub
