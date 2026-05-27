Option Explicit
'*************************************************************************************
'模組名稱: PivotMapChartExample
'功能說明: 以樞紐分析表資料建立地區填色地圖圖表（Excel 2019/365 的 Filled Map Chart）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub CreatePivotMapChart()
    ' xlFilledMap 圖表類型常數（Excel 2019 / Microsoft 365 支援）
    Const xlFilledMap As Long = 107

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chtObj As ChartObject
    Dim cht As Chart
    Dim i As Long

    On Error GoTo ErrHandler

    ' 清除舊工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("地圖資料").Delete
    ThisWorkbook.Sheets("地圖樞紐").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    ' 建立示範資料工作表（英文城市名稱供地圖辨識）
    Set wsData = ThisWorkbook.Sheets.Add
    wsData.Name = "地圖資料"

    wsData.Cells(1, 1).Value = "City"
    wsData.Cells(1, 2).Value = "Sales"
    wsData.Rows(1).Font.Bold = True

    Dim regions(1 To 10, 1 To 2) As Variant
    regions(1, 1) = "Taipei":      regions(1, 2) = 850000
    regions(2, 1) = "New Taipei":  regions(2, 2) = 720000
    regions(3, 1) = "Taoyuan":     regions(3, 2) = 540000
    regions(4, 1) = "Taichung":    regions(4, 2) = 610000
    regions(5, 1) = "Tainan":      regions(5, 2) = 430000
    regions(6, 1) = "Kaohsiung":   regions(6, 2) = 580000
    regions(7, 1) = "Hsinchu":     regions(7, 2) = 320000
    regions(8, 1) = "Keelung":     regions(8, 2) = 210000
    regions(9, 1) = "Chiayi":      regions(9, 2) = 180000
    regions(10, 1) = "Hualien":    regions(10, 2) = 150000

    For i = 1 To 10
        wsData.Cells(i + 1, 1).Value = regions(i, 1)
        wsData.Cells(i + 1, 2).Value = regions(i, 2)
    Next i
    wsData.Columns("A:B").AutoFit

    ' 建立樞紐分析表工作表
    Set wsPivot = ThisWorkbook.Sheets.Add(After:=wsData)
    wsPivot.Name = "地圖樞紐"

    ' 建立樞紐快取
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1:B11"))

    ' 建立樞紐分析表
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="地圖樞紐表")

    With pt.PivotFields("City")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields("Sales")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0"
        .Name = "加總 - Sales"
    End With

    ' 建立圖表物件
    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Cells(1, 5).Left, _
        Top:=wsPivot.Cells(1, 5).Top, _
        Width:=500, Height:=340)

    Set cht = chtObj.Chart

    ' 嘗試建立地圖圖表（需 Excel 2019 以上）
    On Error Resume Next
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlFilledMap

    If Err.Number <> 0 Then
        ' 版本不支援時改用長條圖替代
        Err.Clear
        cht.ChartType = xlBarClustered
        cht.SetSourceData Source:=wsData.Range("A1:B11")
        cht.HasTitle = True
        cht.ChartTitle.Text = "地區銷售分佈（長條圖替代）"
        MsgBox "您的 Excel 版本不支援地圖圖表，已改用長條圖顯示。", vbInformation, "提示"
    Else
        cht.HasTitle = True
        cht.ChartTitle.Text = "地區銷售填色地圖"
    End If
    On Error GoTo ErrHandler

    wsPivot.Activate
    MsgBox "地圖圖表建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
