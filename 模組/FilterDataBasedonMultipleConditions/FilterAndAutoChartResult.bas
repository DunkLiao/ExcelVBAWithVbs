Attribute VB_Name = "FilterAndAutoChartResult"
Option Explicit
'*************************************************************************************
'模組名稱: FilterAndAutoChartResult
'功能說明: 依多重條件篩選資料後，自動產生篩選結果摘要圖表（圓餅圖與長條圖）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestFilterAndAutoChartResult()
    Call FilterAndAutoChartResult("電子產品", 200)
End Sub

Sub FilterAndAutoChartResult(ByVal category As String, ByVal minAmount As Double)
    Dim wsSource As Worksheet
    Dim wsResult As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim resultRow As Long
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim cat As String
    Dim amt As Double
    
    sheetName = "篩選圖表來源"
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        Set wsSource = ThisWorkbook.Worksheets.Add
        wsSource.Name = sheetName
    End If
    
    wsSource.Cells.Clear
    Call FillFilterChartData(wsSource)
    
    On Error Resume Next
    Set wsResult = ThisWorkbook.Worksheets("篩選圖表結果")
    If Not wsResult Is Nothing Then wsResult.Delete
    On Error GoTo 0
    
    Set wsResult = ThisWorkbook.Worksheets.Add
    wsResult.Name = "篩選圖表結果"
    
    wsResult.Range("A1").Value = "產品"
    wsResult.Range("B1").Value = "類別"
    wsResult.Range("C1").Value = "金額"
    resultRow = 1
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        cat = CStr(wsSource.Cells(i, 2).Value)
        amt = CDbl(wsSource.Cells(i, 3).Value)
        
        If cat = category And amt >= minAmount Then
            resultRow = resultRow + 1
            wsSource.Rows(i).Copy wsResult.Rows(resultRow)
        End If
    Next i
    
    If resultRow > 1 Then
        Set chartObj = wsResult.ChartObjects.Add( _
            Left:=wsResult.Range("F1").Left, _
            Top:=wsResult.Range("F1").Top, _
            Width:=480, _
            Height:=320)
        
        Set cht = chartObj.Chart
        cht.ChartType = xlPie
        
        cht.SetSourceData Source:=wsResult.Range("A1:C" & resultRow)
        cht.SeriesCollection(1).XValues = wsResult.Range("A2:A" & resultRow)
        cht.SeriesCollection(1).Values = wsResult.Range("C2:C" & resultRow)
        
        cht.HasTitle = True
        cht.ChartTitle.Text = category & "篩選結果"
        
        cht.ApplyDataLabels
        cht.SeriesCollection(1).DataLabels.ShowPercentage = True
        cht.SeriesCollection(1).DataLabels.ShowCategoryName = True
    End If
    
    wsResult.Columns("A:C").AutoFit
    wsResult.Activate
    
    MsgBox "多重條件篩選並自動產生圖表完成！" & vbCrLf & _
           "篩選條件: 類別=" & category & ", 金額>=" & minAmount & vbCrLf & _
           "符合筆數: " & resultRow - 1 & " 筆", vbInformation, "完成"
End Sub

Private Sub FillFilterChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "類別"
    ws.Range("C1").Value = "金額"
    
    ws.Range("A2").Value = "筆記型電腦"
    ws.Range("B2").Value = "電子產品"
    ws.Range("C2").Value = 450
    ws.Range("A3").Value = "智慧手機"
    ws.Range("B3").Value = "電子產品"
    ws.Range("C3").Value = 380
    ws.Range("A4").Value = "平板電腦"
    ws.Range("B4").Value = "電子產品"
    ws.Range("C4").Value = 150
    ws.Range("A5").Value = "辦公桌"
    ws.Range("B5").Value = "家具"
    ws.Range("C5").Value = 280
    ws.Range("A6").Value = "印表機"
    ws.Range("B6").Value = "電子產品"
    ws.Range("C6").Value = 220
    ws.Range("A7").Value = "辦公椅"
    ws.Range("B7").Value = "家具"
    ws.Range("C7").Value = 180
    ws.Range("A8").Value = "藍牙耳機"
    ws.Range("B8").Value = "電子產品"
    ws.Range("C8").Value = 320
    ws.Range("A9").Value = "書櫃"
    ws.Range("B9").Value = "家具"
    ws.Range("C9").Value = 250
    
    ws.Columns("A:C").AutoFit
End Sub
