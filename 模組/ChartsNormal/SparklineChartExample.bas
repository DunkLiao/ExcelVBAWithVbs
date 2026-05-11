Attribute VB_Name = "SparklineChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: SparklineChartExample
'功能說明: 在 Excel 儲存格中建立折線走勢圖與直條走勢圖的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口
Sub TestSparklineChart()
    Call CreateSparklines("走勢圖範例")
End Sub

' 建立走勢圖
' sheetName: 要建立走勢圖的工作表名稱
Sub CreateSparklines(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim sGroup As SparklineGroup

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    Call FillSparklineData(ws)

    ' 在 H2:H6 建立折線走勢圖
    Set rng = ws.Range("H2:H6")
    Set sGroup = ws.SparklineGroups.Add( _
        Type:=xlSparkLine, _
        SourceData:="B2:G6")
    sGroup.Location = rng

    With sGroup
        .SeriesColor.Color = RGB(70, 130, 180)
        .Points.Highpoint.Visible = True
        .Points.Highpoint.Color.Color = RGB(255, 0, 0)
        .Points.Lowpoint.Visible = True
        .Points.Lowpoint.Color.Color = RGB(0, 128, 0)
    End With

    ' 在 H8:H12 建立直條走勢圖
    Set rng = ws.Range("H8:H12")
    Set sGroup = ws.SparklineGroups.Add( _
        Type:=xlSparkColumn, _
        SourceData:="B8:G12")
    sGroup.Location = rng

    ws.Columns("A:H").AutoFit
    MsgBox "走勢圖已建立完成！", vbInformation, "完成"
End Sub

' 填入走勢圖範例資料
Private Sub FillSparklineData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "一月"
    ws.Range("C1").Value = "二月"
    ws.Range("D1").Value = "三月"
    ws.Range("E1").Value = "四月"
    ws.Range("F1").Value = "五月"
    ws.Range("G1").Value = "六月"
    ws.Range("H1").Value = "走勢"
    ws.Range("A1:H1").Font.Bold = True

    ws.Range("A2").Value = "產品A"
    ws.Range("B2:G2").Value = Array(120, 135, 108, 145, 162, 155)
    ws.Range("A3").Value = "產品B"
    ws.Range("B3:G3").Value = Array(85, 92, 110, 98, 87, 105)
    ws.Range("A4").Value = "產品C"
    ws.Range("B4:G4").Value = Array(200, 180, 220, 195, 210, 230)
    ws.Range("A5").Value = "產品D"
    ws.Range("B5:G5").Value = Array(60, 75, 68, 80, 72, 90)
    ws.Range("A6").Value = "產品E"
    ws.Range("B6:G6").Value = Array(150, 160, 155, 170, 165, 180)

    ws.Range("A7").Value = "部門"
    ws.Range("B7:G7").Value = Array("一月", "二月", "三月", "四月", "五月", "六月")
    ws.Range("H7").Value = "走勢"
    ws.Range("A7:H7").Font.Bold = True

    ws.Range("A8").Value = "業務部"
    ws.Range("B8:G8").Value = Array(320, 280, 350, 310, 370, 400)
    ws.Range("A9").Value = "行銷部"
    ws.Range("B9:G9").Value = Array(180, 200, 175, 220, 195, 215)
    ws.Range("A10").Value = "研發部"
    ws.Range("B10:G10").Value = Array(450, 470, 440, 490, 510, 480)
    ws.Range("A11").Value = "人資部"
    ws.Range("B11:G11").Value = Array(90, 95, 88, 100, 92, 105)
    ws.Range("A12").Value = "財務部"
    ws.Range("B12:G12").Value = Array(260, 280, 270, 300, 290, 310)
End Sub