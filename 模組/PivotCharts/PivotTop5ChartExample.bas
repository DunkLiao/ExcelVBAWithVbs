Attribute VB_Name = "PivotTop5ChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotTop5ChartExample
'功能說明: 建立顯示銷售前五名客戶的樞紐分析圖，使用值篩選功能
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

Sub TestPivotTop5Chart()
    Call CreatePivotTop5Chart
End Sub

Sub CreatePivotTop5Chart()
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim pf      As PivotField
    Dim chtObj  As ChartObject
    Dim cht     As Chart

    On Error GoTo ErrHandler

    Set wsData  = Top5GetOrCreateWs("客戶銷售資料")
    Set wsPivot = Top5GetOrCreateWs("前五名樞紐")

    Call FillTop5Data(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="客戶銷售樞紐")

    With pt
        Set pf = .PivotFields("客戶名稱")
        pf.Orientation = xlRowField
        pf.Position = 1
        .AddDataField .PivotFields("購買金額"), "購買金額合計", xlSum
        .DataFields(1).NumberFormat = "#,##0"
    End With

    '套用前五名值篩選
    pt.PivotFields("客戶名稱").PivotFilters.Add2 _
        Type:=xlTopCount, _
        DataField:=pt.DataFields(1), _
        Value1:=5

    '依金額降冪排序
    pt.PivotFields("客戶名稱").AutoSort xlDescending, "購買金額合計"

    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("D3").Left, _
        Top:=wsPivot.Range("D3").Top, _
        Width:=420, Height:=280)

    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlBarClustered
    cht.HasTitle = True
    cht.ChartTitle.Text = "銷售前五名客戶排行圖"
    cht.HasLegend = False
    cht.SeriesCollection(1).HasDataLabels = True
    cht.SeriesCollection(1).DataLabels.NumberFormat = "#,##0"

    wsPivot.Activate
    MsgBox "前五名客戶樞紐圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立前五名樞紐圖失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillTop5Data(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "客戶名稱"
    ws.Range("B1").Value = "購買金額"
    ws.Range("A2").Value = "台灣科技股份有限公司"
    ws.Range("B2").Value = 580000
    ws.Range("A3").Value = "創新電子有限公司"
    ws.Range("B3").Value = 420000
    ws.Range("A4").Value = "全球貿易股份有限公司"
    ws.Range("B4").Value = 750000
    ws.Range("A5").Value = "信義實業有限公司"
    ws.Range("B5").Value = 310000
    ws.Range("A6").Value = "大安科技股份有限公司"
    ws.Range("B6").Value = 890000
    ws.Range("A7").Value = "中山商貿有限公司"
    ws.Range("B7").Value = 225000
    ws.Range("A8").Value = "松山企業股份有限公司"
    ws.Range("B8").Value = 640000
    ws.Range("A9").Value = "內湖科學園區有限公司"
    ws.Range("B9").Value = 480000
    ws.Range("A10").Value = "南港電子股份有限公司"
    ws.Range("B10").Value = 390000
    ws.Range("A11").Value = "士林國際有限公司"
    ws.Range("B11").Value = 175000
    ws.Columns("A:B").AutoFit
End Sub

Private Function Top5GetOrCreateWs(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set Top5GetOrCreateWs = ws
End Function
