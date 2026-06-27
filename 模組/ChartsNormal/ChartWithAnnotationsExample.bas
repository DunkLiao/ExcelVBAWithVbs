Attribute VB_Name = "ChartWithAnnotationsExample"
Option Explicit
'*************************************************************************************
'模組名稱: ChartWithAnnotationsExample
'功能說明: 在 Excel 圖表中加入文字標註與箭頭說明線的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestChartWithAnnotations()
    Call CreateChartWithAnnotations
End Sub

Sub CreateChartWithAnnotations()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range

    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets("圖表標註範例")
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "圖表標註範例"

    ' 填入範例資料
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "營業額"
    ws.Range("A2").Value = "1月"
    ws.Range("B2").Value = 1500
    ws.Range("A3").Value = "2月"
    ws.Range("B3").Value = 1800
    ws.Range("A4").Value = "3月"
    ws.Range("B4").Value = 1200
    ws.Range("A5").Value = "4月"
    ws.Range("B5").Value = 2200
    ws.Range("A6").Value = "5月"
    ws.Range("B6").Value = 1900
    ws.Range("A1:B1").Font.Bold = True

    Set dataRange = ws.Range("A1:B6")

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=500, _
        Height:=350)

    Set cht = chartObj.Chart
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlLineMarkers
    cht.HasTitle = True
    cht.ChartTitle.Text = "上半年營業額趨勢"
    cht.ChartStyle = 2

    ' 加入最高點標註
    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, 350, 80, 120, 30)
    shp.TextFrame.Characters.Text = "最高點"
    shp.TextFrame.Characters.Font.Size = 10
    shp.TextFrame.Characters.Font.Color = RGB(192, 0, 0)
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse

    ' 加入箭頭
    Dim arrowShape As Shape
    Set arrowShape = cht.Shapes.AddConnector( _
        msoConnectorStraight, 320, 120, 370, 100)
    arrowShape.Line.ForeColor.RGB = RGB(192, 0, 0)
    arrowShape.Line.EndArrowheadStyle = msoArrowheadTriangle

    ws.Columns("A:B").AutoFit
    MsgBox "包含標註的圖表已建立完成！", vbInformation, "完成"
End Sub
