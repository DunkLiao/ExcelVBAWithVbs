Option Explicit
Attribute VB_Name = "CylinderChartExample"
'*************************************************************************************
'模組名稱: 圓柱圖範例
'功能說明: 在 Excel 中建立 3D 直條（圓柱）圖的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub TestCylinderChart()
    Call CreateCylinderChart("圓柱圖範例")
End Sub

Sub CreateCylinderChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheetCylinder(sheetName)
    ws.Cells.Clear

    Call FillCylinderChartData(ws)
    Set dataRange = ws.Range("A1:D5")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("F1").Left, _
        Top:=ws.Range("F1").Top, _
        Width:=480, _
        Height:=320)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xl3DColumn

    cht.HasTitle = True
    cht.ChartTitle.Text = "各區域季度銷售量（3D直條圖）"

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "季度"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售量"
    End With

    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    MsgBox "圓柱圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立圓柱圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetOrCreateWorksheetCylinder(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheetCylinder = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheetCylinder Is Nothing Then
        Set GetOrCreateWorksheetCylinder = ThisWorkbook.Worksheets.Add
        GetOrCreateWorksheetCylinder.Name = sheetName
    End If
End Function

Private Sub FillCylinderChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "季度"
    ws.Range("B1").Value = "北區"
    ws.Range("C1").Value = "中區"
    ws.Range("D1").Value = "南區"

    ws.Range("A2").Value = "Q1"
    ws.Range("B2").Value = 320
    ws.Range("C2").Value = 280
    ws.Range("D2").Value = 250

    ws.Range("A3").Value = "Q2"
    ws.Range("B3").Value = 350
    ws.Range("C3").Value = 310
    ws.Range("D3").Value = 275

    ws.Range("A4").Value = "Q3"
    ws.Range("B4").Value = 410
    ws.Range("C4").Value = 360
    ws.Range("D4").Value = 300

    ws.Range("A5").Value = "Q4"
    ws.Range("B5").Value = 480
    ws.Range("C5").Value = 420
    ws.Range("D5").Value = 350

    ws.Columns("A:D").AutoFit
End Sub
