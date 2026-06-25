Option Explicit
Attribute VB_Name = "Pivot3DColumnChartExample"
'*************************************************************************************
'模組名稱: Pivot3DColumnChartExample
'功能說明: 示範如何從樞紐分析表建立3D立體直條圖
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestPivot3DColumnChart()
    Call CreatePivot3DColumnChart
End Sub

' 從樞紐分析表建立3D立體直條圖
Sub CreatePivot3DColumnChart()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheet("樞紐3D圖表範例")
    ws.Cells.Clear

    ' 建立範例資料
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "季別"
    ws.Range("C1").Value = "銷售額"
    ws.Range("A2").Value = "北部"
    ws.Range("B2").Value = "Q1"
    ws.Range("C2").Value = 5000
    ws.Range("A3").Value = "北部"
    ws.Range("B3").Value = "Q2"
    ws.Range("C3").Value = 6000
    ws.Range("A4").Value = "中部"
    ws.Range("B4").Value = "Q1"
    ws.Range("C4").Value = 4000
    ws.Range("A5").Value = "中部"
    ws.Range("B5").Value = "Q2"
    ws.Range("C5").Value = 4500
    ws.Range("A6").Value = "南部"
    ws.Range("B6").Value = "Q1"
    ws.Range("C6").Value = 3500
    ws.Range("A7").Value = "南部"
    ws.Range("B7").Value = "Q2"
    ws.Range("C7").Value = 4200

    Set dataRange = ws.Range("A1:C7")

    ' 建立樞紐分析表
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=ws.Range("E1"), _
        TableName:="地區樞紐")

    With pt
        .PivotFields("地區").Orientation = xlRowField
        .PivotFields("季別").Orientation = xlColumnField
        .PivotFields("銷售額").Orientation = xlDataField
    End With

    ' 從樞紐建立3D立體直條圖
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E10").Left, _
        Top:=ws.Range("E10").Top, _
        Width:=550, _
        Height:=380)
    Set cht = chartObj.Chart

    ' 設定圖表類型為3D直條圖
    cht.ChartType = xl3DColumnClustered
    cht.SetSourceData Source:=pt.TableRange1

    cht.HasTitle = True
    cht.ChartTitle.Text = "各地區季別銷售額3D圖"

    ' 設定3D視角
    If cht.Has3DEffect Then
        cht.Rotation = 30
        cht.Elevation = 20
    End If

    MsgBox "已建立樞紐3D立體直條圖。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
