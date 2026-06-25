Option Explicit
Attribute VB_Name = "ChartDataLabelExample"
'*************************************************************************************
'模組名稱: ChartDataLabelExample
'功能說明: 示範如何自訂圖表資料標籤，包含數值、百分比與類別名稱格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestChartDataLabel()
    Call CreateChartWithDataLabel("資料標籤範例")
End Sub

' 建立含自訂資料標籤的圖表
Sub CreateChartWithDataLabel(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range
    Dim i As Long

    Set ws = GetOrCreateWorksheet(sheetName)
    ws.Cells.Clear

    ' 填入範例資料
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "營業額"
    ws.Range("A2").Value = "一月"
    ws.Range("B2").Value = 120
    ws.Range("A3").Value = "二月"
    ws.Range("B3").Value = 150
    ws.Range("A4").Value = "三月"
    ws.Range("B4").Value = 180
    ws.Range("A5").Value = "四月"
    ws.Range("B5").Value = 200
    ws.Range("A6").Value = "五月"
    ws.Range("B6").Value = 170

    Set dataRange = ws.Range("A1:B6")

    ' 建立長條圖
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("D1").Left, _
        Top:=ws.Range("D1").Top, _
        Width:=500, _
        Height:=350)
    Set cht = chartObj.Chart
    cht.ChartType = xlColumnClustered
    cht.SetSourceData Source:=dataRange

    cht.HasTitle = True
    cht.ChartTitle.Text = "各月營業額"

    ' 在數列上新增資料標籤
    If cht.SeriesCollection.Count > 0 Then
        With cht.SeriesCollection(1)
            .HasDataLabels = True
            .DataLabels.ShowValue = True
            .DataLabels.ShowCategoryName = True
            .DataLabels.Font.Size = 10
            .DataLabels.Font.Bold = True
        End With
    End If

    MsgBox "已建立含自訂資料標籤的圖表。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
