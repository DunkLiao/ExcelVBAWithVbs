Attribute VB_Name = "ChartWithDynamicTitleExample"
Option Explicit
'*************************************************************************************
'模組名稱: ChartWithDynamicTitleExample
'功能說明: 建立圖表並將圖表標題動態連結至儲存格數值，修改儲存格後標題自動更新
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestChartWithDynamicTitle()
    Call CreateChartWithDynamicTitle
End Sub

Sub CreateChartWithDynamicTitle()
    Dim ws As Worksheet
    Dim chtObj As ChartObject
    Dim cht As Chart
    Dim rngData As Range
    Dim rngTitle As Range
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim wsName As String
    wsName = "動態標題圖表"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName
    
    ws.Range("A1").Value = "圖表標題控制"
    ws.Range("B1").Value = "2026年各季銷售數據"
    ws.Range("A1:B1").Font.Bold = True
    Set rngTitle = ws.Range("B1")
    
    ws.Range("A3").Value = "季度"
    ws.Range("B3").Value = "銷售額"
    ws.Range("A4").Value = "Q1"
    ws.Range("B4").Value = 150000
    ws.Range("A5").Value = "Q2"
    ws.Range("B5").Value = 220000
    ws.Range("A6").Value = "Q3"
    ws.Range("B6").Value = 180000
    ws.Range("A7").Value = "Q4"
    ws.Range("B7").Value = 250000
    
    ws.Range("B4:B7").NumberFormat = "#,##0"
    
    Set rngData = ws.Range("A3:B7")
    
    Set chtObj = ws.ChartObjects.Add(Left:=10, Top:=180, Width:=480, Height:=300)
    Set cht = chtObj.Chart
    cht.SetSourceData Source:=rngData
    cht.ChartType = xlColumnClustered
    
    cht.HasTitle = True
    cht.ChartTitle.Text = "=" & ws.Name & "!" & rngTitle.Address
    
    cht.HasLegend = False
    
    ws.Columns("A:B").AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "動態標題圖表建立完成！" & vbCrLf & _
           "修改儲存格 B1 的內容，圖表標題會自動更新。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "建立動態標題圖表時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
