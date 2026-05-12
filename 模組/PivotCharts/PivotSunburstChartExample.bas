Option Explicit
Attribute VB_Name = "PivotSunburstChartExample"
'*************************************************************************************
'模組名稱: PivotSunburstChartExample
'功能說明: 建立樞紐分析表並從中產生旭日圖（Sunburst Chart）呈現階層式資料結構
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub CreatePivotSunburstChartExample()
    Dim wbk As Workbook
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim cht As ChartObject
    Dim dataArr(1 To 9, 1 To 3) As Variant

    On Error GoTo ErrHandler

    Set wbk = ThisWorkbook

    Application.DisplayAlerts = False
    On Error Resume Next
    wbk.Sheets("資料來源").Delete
    wbk.Sheets("樞紐旭日圖").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set wsData = wbk.Sheets.Add(After:=wbk.Sheets(wbk.Sheets.Count))
    wsData.Name = "資料來源"
    wsData.Range("A1:C1").Value = Array("大區", "區域", "業績")

    dataArr(1, 1) = "北部" : dataArr(1, 2) = "台北" : dataArr(1, 3) = 180000
    dataArr(2, 1) = "北部" : dataArr(2, 2) = "新北" : dataArr(2, 3) = 140000
    dataArr(3, 1) = "北部" : dataArr(3, 2) = "桃園" : dataArr(3, 3) = 110000
    dataArr(4, 1) = "中部" : dataArr(4, 2) = "台中" : dataArr(4, 3) = 130000
    dataArr(5, 1) = "中部" : dataArr(5, 2) = "彰化" : dataArr(5, 3) = 70000
    dataArr(6, 1) = "中部" : dataArr(6, 2) = "南投" : dataArr(6, 3) = 45000
    dataArr(7, 1) = "南部" : dataArr(7, 2) = "台南" : dataArr(7, 3) = 120000
    dataArr(8, 1) = "南部" : dataArr(8, 2) = "高雄" : dataArr(8, 3) = 160000
    dataArr(9, 1) = "南部" : dataArr(9, 2) = "屏東" : dataArr(9, 3) = 55000

    wsData.Range("A2:C10").Value = dataArr
    wsData.Columns.AutoFit

    Set wsPivot = wbk.Sheets.Add(After:=wbk.Sheets(wbk.Sheets.Count))
    wsPivot.Name = "樞紐旭日圖"

    Set pc = wbk.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1:C10"))

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="PivotSunburst")

    With pt
        .PivotFields("大區").Orientation = xlRowField
        .PivotFields("大區").Position = 1
        .PivotFields("區域").Orientation = xlRowField
        .PivotFields("區域").Position = 2
        .PivotFields("業績").Orientation = xlDataField
        .PivotFields("業績").Function = xlSum
        .PivotFields("業績").NumberFormat = "#,##0"
    End With

    Set cht = wsPivot.ChartObjects.Add(Left:=200, Top:=10, Width:=380, Height:=300)
    With cht.Chart
        .SetSourceData Source:=pt.TableRange1
        .ChartType = xlSunburst
        .HasTitle = True
        .ChartTitle.Text = "區域業績旭日圖"
    End With

    MsgBox "樞紐旭日圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "錯誤：" & Err.Description, vbCritical, "建立樞紐旭日圖失敗"
End Sub