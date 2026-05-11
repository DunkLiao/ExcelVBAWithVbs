Attribute VB_Name = "SunburstChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: SunburstChartExample
'功能說明: 在 Excel 中建立旭日圖 (Sunburst) 的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestSunburstChart()
    Call CreateSunburstChart("旭日圖範例")
End Sub

' 建立旭日圖
' sheetName: 要建立圖表的工作表名稱
Sub CreateSunburstChart(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    Set ws = GetOrCreateWorksheetSunburst(sheetName)
    ws.Cells.Clear

    Call FillSunburstData(ws)
    Set dataRange = ws.Range("A1:C10")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=480, _
        Height:=360)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlSunburst

    cht.HasTitle = True
    cht.ChartTitle.Text = "各區域部門銷售分佈旭日圖"

    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    MsgBox "旭日圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立旭日圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheetSunburst(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheetSunburst = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheetSunburst Is Nothing Then
        Set GetOrCreateWorksheetSunburst = ThisWorkbook.Worksheets.Add
        GetOrCreateWorksheetSunburst.Name = sheetName
    End If
End Function

' 填入旭日圖範例資料（區域 > 部門 > 銷售額）
Private Sub FillSunburstData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "區域"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "銷售額"

    ws.Range("A2").Value = "北區"
    ws.Range("B2").Value = "業務一部"
    ws.Range("C2").Value = 320

    ws.Range("A3").Value = "北區"
    ws.Range("B3").Value = "業務二部"
    ws.Range("C3").Value = 280

    ws.Range("A4").Value = "中區"
    ws.Range("B4").Value = "業務一部"
    ws.Range("C4").Value = 210

    ws.Range("A5").Value = "中區"
    ws.Range("B5").Value = "業務二部"
    ws.Range("C5").Value = 190

    ws.Range("A6").Value = "南區"
    ws.Range("B6").Value = "業務一部"
    ws.Range("C6").Value = 250

    ws.Range("A7").Value = "南區"
    ws.Range("B7").Value = "業務二部"
    ws.Range("C7").Value = 230

    ws.Range("A8").Value = "東區"
    ws.Range("B8").Value = "業務一部"
    ws.Range("C8").Value = 180

    ws.Range("A9").Value = "東區"
    ws.Range("B9").Value = "業務二部"
    ws.Range("C9").Value = 160

    ws.Range("A10").Value = "西區"
    ws.Range("B10").Value = "業務一部"
    ws.Range("C10").Value = 140

    ws.Columns("A:C").AutoFit
End Sub
