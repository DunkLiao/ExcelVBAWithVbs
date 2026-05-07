Attribute VB_Name = "ValueStatsPivot"
Option Explicit
'*************************************************************************************
'模組名稱: ValueStatsPivot
'功能說明: 多種統計值示範
'          在同一樞紐分析表中，對同一資料欄同時顯示
'          加總、平均、最大值、最小值、計數
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestValueStatsPivot()
    Call CreateValueStatsPivot
End Sub

' 建立多統計值樞紐分析表
Sub CreateValueStatsPivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim dfSum As PivotField
    Dim dfAvg As PivotField
    Dim dfMax As PivotField
    Dim dfMin As PivotField
    Dim dfCnt As PivotField

    Set wsData = GetOrCreateSheet(ThisWorkbook, "業務銷售")
    Call FillSalesDetail(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "多統計值樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="多統計值樞紐分析表")

    ' 列欄位：業務員
    With pt.PivotFields("業務員")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 加總
    Set dfSum = pt.AddDataField(pt.PivotFields("金額"), "金額加總", xlSum)
    dfSum.NumberFormat = "#,##0"

    ' 平均
    Set dfAvg = pt.AddDataField(pt.PivotFields("金額"), "金額平均", xlAverage)
    dfAvg.NumberFormat = "#,##0.00"

    ' 最大值
    Set dfMax = pt.AddDataField(pt.PivotFields("金額"), "金額最大", xlMax)
    dfMax.NumberFormat = "#,##0"

    ' 最小值
    Set dfMin = pt.AddDataField(pt.PivotFields("金額"), "金額最小", xlMin)
    dfMin.NumberFormat = "#,##0"

    ' 計數
    Set dfCnt = pt.AddDataField(pt.PivotFields("金額"), "筆數", xlCount)

    ' 值欄位排列為欄方向
    pt.DataPivotField.Orientation = xlColumnField

    wsPivot.Columns("A:G").AutoFit
    wsPivot.Activate
    MsgBox "多統計值樞紐分析表已建立！" & Chr(13) & _
           "同時顯示加總、平均、最大、最小、計數。", vbInformation, "完成"
End Sub

' 填入業務銷售明細
Private Sub FillSalesDetail(ByVal ws As Worksheet)
    ws.Range("A1").Value = "業務員"
    ws.Range("B1").Value = "月份"
    ws.Range("C1").Value = "金額"

    ws.Range("A2").Value = "王小明" : ws.Range("B2").Value = "一月" : ws.Range("C2").Value = 32000
    ws.Range("A3").Value = "王小明" : ws.Range("B3").Value = "二月" : ws.Range("C3").Value = 45000
    ws.Range("A4").Value = "王小明" : ws.Range("B4").Value = "三月" : ws.Range("C4").Value = 28000
    ws.Range("A5").Value = "李美華" : ws.Range("B5").Value = "一月" : ws.Range("C5").Value = 51000
    ws.Range("A6").Value = "李美華" : ws.Range("B6").Value = "二月" : ws.Range("C6").Value = 39000
    ws.Range("A7").Value = "李美華" : ws.Range("B7").Value = "三月" : ws.Range("C7").Value = 62000
    ws.Range("A8").Value = "張志偉" : ws.Range("B8").Value = "一月" : ws.Range("C8").Value = 18000
    ws.Range("A9").Value = "張志偉" : ws.Range("B9").Value = "二月" : ws.Range("C9").Value = 27000
    ws.Range("A10").Value = "張志偉" : ws.Range("B10").Value = "三月" : ws.Range("C10").Value = 33000

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
