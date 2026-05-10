Attribute VB_Name = "WaterfallChartExample"
Option Explicit

' ============================================================
' 模組名稱：WaterfallChartExample
' 功能說明：建立瀑布圖展示資金流入與流出變化
' 適用版本：Excel 2016 以上
' ============================================================

Sub CreateWaterfallChartExample()
    Dim ws          As Worksheet
    Dim chtObj      As ChartObject
    Dim cht         As Chart
    Dim rngData     As Range
    Dim i           As Integer
    
    ' 資料陣列
    Dim arrItems(7)  As String
    Dim arrValues(7) As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    ' 若工作表已存在則刪除
    Dim wsName As String
    wsName = "瀑布圖範例"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    ' 新增工作表
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName
    
    ' 設定標題列
    With ws.Range("A1:B1")
        .Value = Array("項目", "金額")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' 準備範例資料（現金流量）
    arrItems(0) = "期初餘額"  : arrValues(0) = 100000
    arrItems(1) = "銷售收入"  : arrValues(1) = 50000
    arrItems(2) = "其他收入"  : arrValues(2) = 10000
    arrItems(3) = "材料成本"  : arrValues(3) = -20000
    arrItems(4) = "人事費用"  : arrValues(4) = -15000
    arrItems(5) = "管理費用"  : arrValues(5) = -8000
    arrItems(6) = "稅金"      : arrValues(6) = -5000
    arrItems(7) = "期末餘額"  : arrValues(7) = 112000
    
    ' 填入資料
    For i = 0 To 7
        ws.Cells(i + 2, 1).Value = arrItems(i)
        ws.Cells(i + 2, 2).Value = arrValues(i)
    Next i
    
    ' 設定數值格式
    ws.Range("B2:B9").NumberFormat = "#,##0"
    ws.Columns("A:B").AutoFit
    
    ' 定義資料範圍
    Set rngData = ws.Range("A1:B9")
    
    ' 建立圖表物件
    Set chtObj = ws.ChartObjects.Add(Left:=10, Top:=180, Width:=480, Height:=320)
    Set cht = chtObj.Chart
    
    ' 設定圖表來源與類型（xlWaterfall = Excel 2016+）
    cht.SetSourceData Source:=rngData
    cht.ChartType = xlWaterfall
    
    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "現金流量瀑布圖"
    
    ' 設定圖例
    cht.HasLegend = False
    
    Application.ScreenUpdating = True
    MsgBox "瀑布圖建立完成！請查看「" & wsName & "」工作表。", _
           vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "建立瀑布圖時發生錯誤：" & Err.Number & " - " & Err.Description, _
           vbCritical, "錯誤"
End Sub