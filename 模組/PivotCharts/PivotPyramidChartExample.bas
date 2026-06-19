Attribute VB_Name = "PivotPyramidChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotPyramidChartExample
'功能說明: 從樞紐分析表資料建立金字塔圖表（Pyramid Chart）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestPivotPyramidChart()
    Call CreatePivotPyramidChart
End Sub

Sub CreatePivotPyramidChart()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chtObj As ChartObject
    Dim cht As Chart
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim wsName As String
    wsName = "金字塔圖表資料"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    ThisWorkbook.Sheets("金字塔圖表樞紐").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set wsData = ThisWorkbook.Sheets.Add
    wsData.Name = wsName
    
    ' 撰寫範例資料（適合金字塔圖的人口結構資料）
    wsData.Range("A1").Value = "年齡層"
    wsData.Range("B1").Value = "性別"
    wsData.Range("C1").Value = "人數"
    
    Dim ageGroups As Variant
    ageGroups = Array("0-14歲", "15-29歲", "30-44歲", "45-59歲", "60-74歲", "75歲以上")
    Dim genderArr As Variant
    genderArr = Array("男性", "女性")
    Dim popVals As Variant
    popVals = Array(8500, 12000, 15000, 13000, 9500, 5000, _
                     8000, 11500, 14000, 12500, 9000, 4500)
    
    Dim i As Long, j As Long, k As Long
    k = 0
    For i = 0 To 5
        For j = 0 To 1
            wsData.Cells(k + 2, 1).Value = ageGroups(i)
            wsData.Cells(k + 2, 2).Value = genderArr(j)
            wsData.Cells(k + 2, 3).Value = popVals(k)
            k = k + 1
        Next j
    Next i
    
    ' 建立樞紐分析表
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion.Address)
    
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "金字塔圖表樞紐"
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="樞紐_金字塔")
    
    With pt
        .PivotFields("年齡層").Orientation = xlRowField
        .PivotFields("年齡層").Position = 1
        .PivotFields("性別").Orientation = xlColumnField
        .PivotFields("性別").Position = 1
        .PivotFields("人數").Orientation = xlDataField
        .PivotFields("人數").Function = xlSum
        .PivotFields("人數").NumberFormat = "#,##0"
    End With
    
    ' 建立金字塔圖表（使用橫條圖模擬）
    Set chtObj = wsPivot.ChartObjects.Add(Left:=10, Top:=10, Width:=520, Height:=380)
    Set cht = chtObj.Chart
    cht.SetSourceData Source:=pt.TableRange1
    cht.ChartType = xlBarClustered
    
    cht.HasTitle = True
    cht.ChartTitle.Text = "人口年齡結構金字塔圖"
    
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom
    
    wsPivot.Columns.AutoFit
    wsData.Columns.AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "樞紐金字塔圖表建立完成！" & vbCrLf & _
           "請查看「金字塔圖表樞紐」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "建立金字塔圖表時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
