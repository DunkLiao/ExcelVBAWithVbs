Attribute VB_Name = "MultipleConsolidationPivot"
Option Explicit
'*************************************************************************************
'模組名稱: MultipleConsolidationPivot
'功能說明: 將多個來源工作表資料合併後建立樞紐分析表
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口
Sub TestMultipleConsolidationPivot()
    Call CreateMultipleConsolidationPivot
End Sub

' 建立多重來源資料彙整後的樞紐分析表
Sub CreateMultipleConsolidationPivot()
    Dim wsQ1 As Worksheet
    Dim wsQ2 As Worksheet
    Dim wsCombined As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable

    Set wsQ1 = GetOrCreateSheet(ThisWorkbook, "Q1銷售")
    Set wsQ2 = GetOrCreateSheet(ThisWorkbook, "Q2銷售")
    Call FillQ1Data(wsQ1)
    Call FillQ2Data(wsQ2)

    Set wsCombined = GetOrCreateSheet(ThisWorkbook, "合併資料來源")
    Call CombineSourceData(wsQ1, wsQ2, wsCombined)

    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "多重彙總樞紐")

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsCombined.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="多重彙總樞紐")

    With pt.PivotFields("季度")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields("地區")
        .Orientation = xlColumnField
        .Position = 1
    End With

    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum

    wsPivot.Activate
    MsgBox "多重彙總樞紐分析表已建立完成！", vbInformation, "完成"
End Sub

' 合併兩個來源工作表的資料
Private Sub CombineSourceData(ByVal ws1 As Worksheet, ByVal ws2 As Worksheet, _
                               ByVal wsOut As Worksheet)
    Dim lastRow1 As Long
    Dim lastRow2 As Long

    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    ws1.Range("A1").Resize(lastRow1, 3).Copy Destination:=wsOut.Cells(1, 1)

    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    If lastRow2 >= 2 Then
        ws2.Range("A2").Resize(lastRow2 - 1, 3).Copy Destination:=wsOut.Cells(lastRow1 + 1, 1)
    End If

    wsOut.Columns("A:C").AutoFit
End Sub

' 填入 Q1 銷售資料
Private Sub FillQ1Data(ByVal ws As Worksheet)
    ws.Range("A1").Value = "季度"
    ws.Range("B1").Value = "地區"
    ws.Range("C1").Value = "銷售額"
    ws.Range("A2").Value = "Q1"
    ws.Range("B2").Value = "北區"
    ws.Range("C2").Value = 80000
    ws.Range("A3").Value = "Q1"
    ws.Range("B3").Value = "南區"
    ws.Range("C3").Value = 65000
    ws.Range("A4").Value = "Q1"
    ws.Range("B4").Value = "東區"
    ws.Range("C4").Value = 72000
    ws.Columns("A:C").AutoFit
End Sub

' 填入 Q2 銷售資料
Private Sub FillQ2Data(ByVal ws As Worksheet)
    ws.Range("A1").Value = "季度"
    ws.Range("B1").Value = "地區"
    ws.Range("C1").Value = "銷售額"
    ws.Range("A2").Value = "Q2"
    ws.Range("B2").Value = "北區"
    ws.Range("C2").Value = 95000
    ws.Range("A3").Value = "Q2"
    ws.Range("B3").Value = "南區"
    ws.Range("C3").Value = 78000
    ws.Range("A4").Value = "Q2"
    ws.Range("B4").Value = "東區"
    ws.Range("C4").Value = 88000
    ws.Columns("A:C").AutoFit
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
