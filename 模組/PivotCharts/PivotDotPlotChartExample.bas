Option Explicit
Attribute VB_Name = "PivotDotPlotChartExample"
'*************************************************************************************

'模組名稱: PivotDotPlotChartExample

'功能說明: 以樞紐分析表為資料來源，建立點狀圖（Dot Plot / XY 散點圖）

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub CreatePivotDotPlotChartExample()

    Dim ws As Worksheet

    Dim wsPivot As Worksheet

    Dim pt As PivotTable

    Dim pc As PivotCache

    Dim chtObj As ChartObject

    Dim cht As Chart

    Dim dataRange As Range

    Dim lastRow As Long

    Dim i As Integer



    ' 建立範例資料工作表

    On Error Resume Next

    Application.DisplayAlerts = False

    ThisWorkbook.Worksheets("DotPlotData").Delete

    ThisWorkbook.Worksheets("DotPlotPivot").Delete

    Application.DisplayAlerts = True

    On Error GoTo 0



    Set ws = ThisWorkbook.Worksheets.Add

    ws.Name = "DotPlotData"



    ' 填入標題

    ws.Range("A1").Value = "部門"

    ws.Range("B1").Value = "人員"

    ws.Range("C1").Value = "績效分數"



    ' 填入範例資料

    Dim sampleData As Variant

    sampleData = Array( _

        Array("業務部", "張小明", 85), _

        Array("業務部", "李小華", 92), _

        Array("業務部", "王大同", 78), _

        Array("研發部", "陳志偉", 90), _

        Array("研發部", "林佳穎", 88), _

        Array("研發部", "黃建國", 76), _

        Array("行政部", "吳淑芬", 82), _

        Array("行政部", "劉俊傑", 95), _

        Array("行政部", "蔡明珠", 80))



    For i = 0 To UBound(sampleData)

        ws.Cells(i + 2, 1).Value = sampleData(i)(0)

        ws.Cells(i + 2, 2).Value = sampleData(i)(1)

        ws.Cells(i + 2, 3).Value = sampleData(i)(2)

    Next i



    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set dataRange = ws.Range("A1:C" & lastRow)



    ' 建立樞紐快取

    Set pc = ThisWorkbook.PivotCaches.Create( _

        SourceType:=xlDatabase, _

        SourceData:=dataRange)



    ' 建立樞紐工作表

    Set wsPivot = ThisWorkbook.Worksheets.Add(After:=ws)

    wsPivot.Name = "DotPlotPivot"



    Set pt = pc.CreatePivotTable( _

        TableDestination:=wsPivot.Range("A3"), _

        TableName:="DotPlotPT")



    Application.ScreenUpdating = False



    With pt

        .PivotFields("部門").Orientation = xlRowField

        .PivotFields("部門").Position = 1



        Dim fldScore As PivotField

        Set fldScore = .PivotFields("績效分數")

        fldScore.Orientation = xlDataField

        fldScore.Function = xlAverage

        fldScore.NumberFormat = "0.00"

        fldScore.Name = "平均績效"

    End With



    ' 建立點狀圖（XY 散點圖）

    Set chtObj = wsPivot.ChartObjects.Add(Left:=250, Top:=20, Width:=420, Height:=280)

    Set cht = chtObj.Chart



    cht.SetSourceData Source:=pt.TableRange1

    cht.ChartType = xlXYScatter



    With cht

        .HasTitle = True

        .ChartTitle.Text = "各部門平均績效點狀圖"

        .Axes(xlValue).HasTitle = True

        .Axes(xlValue).AxisTitle.Text = "平均績效分數"

        .Axes(xlCategory).HasTitle = True

        .Axes(xlCategory).AxisTitle.Text = "部門"

        .PlotArea.Interior.Color = RGB(245, 245, 245)

        .ChartArea.Border.Color = RGB(150, 150, 150)

    End With



    wsPivot.Columns.AutoFit

    Application.ScreenUpdating = True



    MsgBox "樞紐點狀圖已建立完成！工作表：" & wsPivot.Name, vbInformation, "完成"

End Sub

