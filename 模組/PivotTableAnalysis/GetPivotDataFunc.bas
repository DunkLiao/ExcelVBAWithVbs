Attribute VB_Name = "GetPivotDataFunc"
Option Explicit
'*************************************************************************************
'模組名稱: GetPivotDataFunc
'功能說明: 使用 GetPivotData 函數從樞紐分析表取得特定值
'          示範以 VBA 呼叫 PivotTable.GetPivotData 讀取指定交叉值
'          並將結果輸出至摘要工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestGetPivotDataFunc()
    Call DemoGetPivotData
End Sub

' 建立樞紐分析表並示範 GetPivotData 取值
Sub DemoGetPivotData()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim wsSummary As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim extractVal As Double

    Set wsData = GetOrCreateSheet(ThisWorkbook, "季度銷售")
    Call FillQuarterData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "季度樞紐")
    Set wsSummary = GetOrCreateSheet(ThisWorkbook, "GetPivotData摘要")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="季度樞紐分析表")

    ' 列欄位：地區
    With pt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 欄欄位：季度
    With pt.PivotFields("季度")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' 值欄位：銷售額
    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum

    ' 建立摘要工作表標題
    wsSummary.Range("A1").Value = "GetPivotData 取值示範"
    wsSummary.Range("A1").Font.Bold = True
    wsSummary.Range("A3").Value = "查詢項目"
    wsSummary.Range("B3").Value = "取得值"
    wsSummary.Range("A3:B3").Font.Bold = True

    ' 以 GetPivotData 取出特定交叉值
    On Error Resume Next
    extractVal = pt.GetPivotData("銷售額合計", "地區", "北部", "季度", "Q1")
    wsSummary.Range("A4").Value = "北部 Q1 銷售額"
    wsSummary.Range("B4").Value = extractVal

    extractVal = pt.GetPivotData("銷售額合計", "地區", "南部", "季度", "Q2")
    wsSummary.Range("A5").Value = "南部 Q2 銷售額"
    wsSummary.Range("B5").Value = extractVal

    extractVal = pt.GetPivotData("銷售額合計", "地區", "東部", "季度", "Q3")
    wsSummary.Range("A6").Value = "東部 Q3 銷售額"
    wsSummary.Range("B6").Value = extractVal
    On Error GoTo 0

    wsSummary.Columns("A:B").AutoFit
    wsSummary.Activate
    MsgBox "GetPivotData 示範完成！" & Chr(13) & _
           "已在『GetPivotData摘要』工作表輸出指定交叉值。", vbInformation, "完成"
End Sub

' 填入季度銷售資料
Private Sub FillQuarterData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "季度"
    ws.Range("C1").Value = "銷售額"

    ws.Range("A2").Value = "北部" : ws.Range("B2").Value = "Q1" : ws.Range("C2").Value = 88000
    ws.Range("A3").Value = "北部" : ws.Range("B3").Value = "Q2" : ws.Range("C3").Value = 92000
    ws.Range("A4").Value = "北部" : ws.Range("B4").Value = "Q3" : ws.Range("C4").Value = 76000
    ws.Range("A5").Value = "北部" : ws.Range("B5").Value = "Q4" : ws.Range("C5").Value = 105000
    ws.Range("A6").Value = "南部" : ws.Range("B6").Value = "Q1" : ws.Range("C6").Value = 65000
    ws.Range("A7").Value = "南部" : ws.Range("B7").Value = "Q2" : ws.Range("C7").Value = 78000
    ws.Range("A8").Value = "南部" : ws.Range("B8").Value = "Q3" : ws.Range("C8").Value = 83000
    ws.Range("A9").Value = "南部" : ws.Range("B9").Value = "Q4" : ws.Range("C9").Value = 91000
    ws.Range("A10").Value = "東部" : ws.Range("B10").Value = "Q1" : ws.Range("C10").Value = 42000
    ws.Range("A11").Value = "東部" : ws.Range("B11").Value = "Q2" : ws.Range("C11").Value = 55000
    ws.Range("A12").Value = "東部" : ws.Range("B12").Value = "Q3" : ws.Range("C12").Value = 61000
    ws.Range("A13").Value = "東部" : ws.Range("B13").Value = "Q4" : ws.Range("C13").Value = 48000

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
