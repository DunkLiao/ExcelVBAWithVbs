Attribute VB_Name = "PivotPageFieldChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotPageFieldChartExample
'功能說明: 建立含頁面欄位（報表篩選）的樞紐分析圖，可依業務員切換檢視
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotPageFieldChart()
    Call CreatePivotPageFieldChart
End Sub

Sub CreatePivotPageFieldChart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim chtObj  As ChartObject
    Dim cht     As Chart

    On Error GoTo ErrHandler

    Set wsData  = PageGetOrCreateWs("業務銷售明細")
    Set wsPivot = PageGetOrCreateWs("頁面篩選樞紐")

    Call FillPageFieldData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="業務銷售樞紐")

    With pt
        '業務員設為頁面欄位（報表篩選）
        .PivotFields("業務員").Orientation = xlPageField
        .PivotFields("業務員").Position = 1
        '月份設為列欄位
        .PivotFields("月份").Orientation = xlRowField
        .PivotFields("月份").Position = 1
        .AddDataField .PivotFields("銷售額"), "銷售額合計", xlSum
        .DataFields(1).NumberFormat = "#,##0"
    End With

    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("E5").Left, _
        Top:=wsPivot.Range("E5").Top, _
        Width:=450, Height:=300)

    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlBarClustered
    cht.HasTitle = True
    cht.ChartTitle.Text = "業務員月份銷售額（可依頁面篩選切換）"
    cht.HasLegend = False

    wsPivot.Activate
    MsgBox "含頁面篩選欄位的樞紐圖已建立！" & vbCrLf & _
           "請在 A1 下拉選單切換業務員以更新圖表。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立頁面篩選樞紐圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillPageFieldData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "業務員"
    ws.Range("B1").Value = "月份"
    ws.Range("C1").Value = "銷售額"
    ws.Range("A2").Value = "陳大明"
    ws.Range("B2").Value = "一月"
    ws.Range("C2").Value = 85000
    ws.Range("A3").Value = "陳大明"
    ws.Range("B3").Value = "二月"
    ws.Range("C3").Value = 92000
    ws.Range("A4").Value = "陳大明"
    ws.Range("B4").Value = "三月"
    ws.Range("C4").Value = 108000
    ws.Range("A5").Value = "林小華"
    ws.Range("B5").Value = "一月"
    ws.Range("C5").Value = 110000
    ws.Range("A6").Value = "林小華"
    ws.Range("B6").Value = "二月"
    ws.Range("C6").Value = 125000
    ws.Range("A7").Value = "林小華"
    ws.Range("B7").Value = "三月"
    ws.Range("C7").Value = 140000
    ws.Range("A8").Value = "王志遠"
    ws.Range("B8").Value = "一月"
    ws.Range("C8").Value = 65000
    ws.Range("A9").Value = "王志遠"
    ws.Range("B9").Value = "二月"
    ws.Range("C9").Value = 72000
    ws.Range("A10").Value = "王志遠"
    ws.Range("B10").Value = "三月"
    ws.Range("C10").Value = 80000
    ws.Columns("A:C").AutoFit
End Sub

Private Function PageGetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set PageGetOrCreateWs = ws
End Function
