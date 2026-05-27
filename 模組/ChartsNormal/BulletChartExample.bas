Option Explicit
'*************************************************************************************
'模組名稱: BulletChartExample
'功能說明: 建立子彈圖（Bullet Chart），以叢集橫條組合圖方式顯示實際值與目標值對比
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub CreateBulletChart()
    Dim ws As Worksheet
    Dim cht As Chart
    Dim srs As Series
    Dim chtObj As ChartObject
    Dim i As Long

    On Error GoTo ErrHandler

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("BulletChart").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "BulletChart"

    ' 標題列
    ws.Cells(1, 1).Value = "項目"
    ws.Cells(1, 2).Value = "實際值"
    ws.Cells(1, 3).Value = "目標值"
    ws.Rows(1).Font.Bold = True

    ' 示範資料（5項產品）
    Dim items(1 To 5) As String
    Dim actuals(1 To 5) As Long
    Dim targets(1 To 5) As Long

    items(1) = "產品A": actuals(1) = 75: targets(1) = 80
    items(2) = "產品B": actuals(2) = 60: targets(2) = 70
    items(3) = "產品C": actuals(3) = 90: targets(3) = 85
    items(4) = "產品D": actuals(4) = 45: targets(4) = 65
    items(5) = "產品E": actuals(5) = 82: targets(5) = 75

    For i = 1 To 5
        ws.Cells(i + 1, 1).Value = items(i)
        ws.Cells(i + 1, 2).Value = actuals(i)
        ws.Cells(i + 1, 3).Value = targets(i)
    Next i

    ' 建立橫條圖（以兩個數列模擬子彈圖效果）
    Set chtObj = ws.ChartObjects.Add( _
        Left:=ws.Cells(8, 1).Left, Top:=ws.Cells(8, 1).Top, _
        Width:=480, Height:=280)
    Set cht = chtObj.Chart
    cht.ChartType = xlBarClustered
    cht.SetSourceData Source:=ws.Range("A1:C6")
    cht.HasTitle = True
    cht.ChartTitle.Text = "子彈圖 - 實際值 vs 目標值"

    ' 實際值數列設定（藍色）
    Set srs = cht.SeriesCollection(1)
    srs.Name = "實際值"
    srs.Interior.Color = RGB(70, 130, 180)

    ' 目標值數列設定（橘紅色）
    Set srs = cht.SeriesCollection(2)
    srs.Name = "目標值"
    srs.Interior.Color = RGB(255, 120, 60)

    cht.PlotArea.Interior.Color = RGB(245, 245, 245)
    ws.Columns("A:C").AutoFit

    MsgBox "子彈圖建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
