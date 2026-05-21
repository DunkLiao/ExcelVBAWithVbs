Option Explicit
Attribute VB_Name = "PivotStepLineChartExample"
'*************************************************************************************
'模組名稱: PivotStepLineChartExample
'功能說明: 以樞紐分析表資料為基礎，建立階梯折線圖呈現數值累計變化
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestPivotStepLineChart()
    Call CreatePivotStepLineChart(ThisWorkbook)
End Sub

Sub CreatePivotStepLineChart(ByVal wb As Workbook)
    On Error GoTo ErrorHandler

    Dim dataWs As Worksheet
    Set dataWs = GetOrCreatePSLSheet(wb, "階梯圖資料來源")
    dataWs.Cells.Clear
    Call FillStepLineData(dataWs)

    Dim pivotWs As Worksheet
    Set pivotWs = GetOrCreatePSLSheet(wb, "階梯折線圖")
    pivotWs.Cells.Clear

    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataWs.Range("A1").CurrentRegion)

    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="階梯圖樞紐")

    pt.PivotFields("季度").Orientation = xlRowField
    pt.PivotFields("季度").Position = 1
    pt.PivotFields("累計業績").Orientation = xlDataField

    Dim chartObj As ChartObject
    Set chartObj = pivotWs.ChartObjects.Add( _
        Left:=pivotWs.Range("F3").Left, _
        Top:=pivotWs.Range("F3").Top, _
        Width:=480, _
        Height:=300)

    Dim cht As Chart
    Set cht = chartObj.Chart
    cht.SetSourceData Source:=pivotWs.Range("A3").CurrentRegion
    cht.ChartType = xlLine

    cht.HasTitle = True
    cht.ChartTitle.Text = "各季度累計業績階梯圖"
    cht.HasLegend = True

    MsgBox "樞紐階梯折線圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立圖表時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillStepLineData(ByVal ws As Worksheet)
    ws.Range("A1:B1").Value = Array("季度", "累計業績")
    ws.Range("A2:B2").Value = Array("Q1", 120000)
    ws.Range("A3:B3").Value = Array("Q2", 250000)
    ws.Range("A4:B4").Value = Array("Q3", 410000)
    ws.Range("A5:B5").Value = Array("Q4", 620000)
    ws.Columns.AutoFit
End Sub

Private Function GetOrCreatePSLSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreatePSLSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreatePSLSheet Is Nothing Then
        Set GetOrCreatePSLSheet = wb.Worksheets.Add
        GetOrCreatePSLSheet.Name = sheetName
    End If
End Function
