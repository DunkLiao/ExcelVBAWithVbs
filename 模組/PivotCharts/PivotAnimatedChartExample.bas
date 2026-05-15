Attribute VB_Name = "PivotAnimatedChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotAnimatedChartExample
'功能說明: 建立樞紐分析表與圖表並以一秒間隔切換圖表類型
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunPivotAnimatedChartExample()
    On Error GoTo ErrorHandler

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim chartTypes As Variant
    Dim chartTitles As Variant
    Dim i As Long

    Set wsData = GetOrCreateAnimatedSheet("樞紐圖表資料")
    Set wsPivot = GetOrCreateAnimatedSheet("樞紐圖表動畫")

    wsData.Cells.Clear
    wsPivot.Cells.Clear

    Call FillAnimatedPivotData(wsData)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1, External:=True))

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="PivotAnimatedChart")

    With pt
        .PivotFields("地區").Orientation = xlRowField
        .PivotFields("地區").Position = 1
        .AddDataField .PivotFields("銷售額"), "銷售額總和", xlSum
    End With

    Set chartObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("E2").Left, _
        Top:=wsPivot.Range("E2").Top, _
        Width:=420, _
        Height:=280)

    chartObj.Chart.SetSourceData Source:=pt.TableRange1
    chartObj.Chart.HasTitle = True
    chartObj.Chart.HasLegend = False

    chartTypes = Array(xlColumnClustered, xlLineMarkers, xlPie)
    chartTitles = Array("樞紐直條圖", "樞紐折線圖", "樞紐圓餅圖")

    For i = LBound(chartTypes) To UBound(chartTypes)
        chartObj.Chart.ChartType = chartTypes(i)
        chartObj.Chart.ChartTitle.Text = chartTitles(i)
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next i

    wsPivot.Columns.AutoFit
    MsgBox "樞紐圖表動畫切換已完成。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立樞紐圖表時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillAnimatedPivotData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("日期", "地區", "產品", "銷售額")
    ws.Range("A2:D2").Value = Array(DateSerial(2026, 1, 5), "北區", "A 產品", 12500)
    ws.Range("A3:D3").Value = Array(DateSerial(2026, 1, 8), "中區", "B 產品", 16800)
    ws.Range("A4:D4").Value = Array(DateSerial(2026, 1, 12), "南區", "A 產品", 14200)
    ws.Range("A5:D5").Value = Array(DateSerial(2026, 1, 15), "東區", "C 產品", 9800)
    ws.Range("A6:D6").Value = Array(DateSerial(2026, 2, 2), "北區", "B 產品", 18700)
    ws.Range("A7:D7").Value = Array(DateSerial(2026, 2, 9), "中區", "C 產品", 13300)
    ws.Range("A8:D8").Value = Array(DateSerial(2026, 2, 14), "南區", "B 產品", 15400)
    ws.Range("A9:D9").Value = Array(DateSerial(2026, 2, 18), "東區", "A 產品", 12100)
    ws.Range("A2:A9").NumberFormat = "yyyy/mm/dd"
    ws.Columns("A:D").AutoFit
End Sub

Private Function GetOrCreateAnimatedSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateAnimatedSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateAnimatedSheet Is Nothing Then
        Set GetOrCreateAnimatedSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateAnimatedSheet.Name = sheetName
    End If
End Function
