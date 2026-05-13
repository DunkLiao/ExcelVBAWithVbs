Attribute VB_Name = "LinkedSlicerPivot"
Option Explicit
'*************************************************************************************
'模組名稱: 連結切片器樞紐分析
'功能說明: 建立兩個樞紐分析表，並以同一個切片器同時控制兩者
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub CreateLinkedSlicerPivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc1 As PivotCache
    Dim pt1 As PivotTable
    Dim pt2 As PivotTable
    Dim sc As SlicerCache
    Dim slicerObj As Slicer

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("切片器資料").Delete
    ThisWorkbook.Worksheets("連結切片器").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsData = ThisWorkbook.Worksheets.Add
    wsData.Name = "切片器資料"
    Call FillLinkedSlicerData(wsData)

    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "連結切片器"

    Set pc1 = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt1 = pc1.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="PT_Sales")

    With pt1
        .PivotFields("地區").Orientation = xlRowField
        .PivotFields("地區").Position = 1
        .AddDataField .PivotFields("銷售額"), "銷售額合計", xlSum
    End With

    Set pt2 = pc1.CreatePivotTable( _
        TableDestination:=wsPivot.Range("E3"), _
        TableName:="PT_Qty")

    With pt2
        .PivotFields("地區").Orientation = xlRowField
        .PivotFields("地區").Position = 1
        .AddDataField .PivotFields("數量"), "數量合計", xlSum
    End With

    Set sc = ThisWorkbook.SlicerCaches.Add2(pt1, "地區")
    Set slicerObj = sc.Slicers.Add(wsPivot, , "地區切片器", "地區", _
        wsPivot.Range("I3").Top, wsPivot.Range("I3").Left, 120, 200)

    sc.PivotTables.AddPivotTable pt2
    wsPivot.Columns.AutoFit

    MsgBox "連結切片器樞紐分析表已建立完成！" & vbCrLf & _
           "切換切片器可同時篩選兩個樞紐分析表。", vbInformation, "完成"
End Sub

Private Sub FillLinkedSlicerData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "產品"
    ws.Range("C1").Value = "銷售額"
    ws.Range("D1").Value = "數量"

    ws.Range("A2").Value = "北區" : ws.Range("B2").Value = "A" : ws.Range("C2").Value = 500 : ws.Range("D2").Value = 10
    ws.Range("A3").Value = "南區" : ws.Range("B3").Value = "B" : ws.Range("C3").Value = 300 : ws.Range("D3").Value = 6
    ws.Range("A4").Value = "北區" : ws.Range("B4").Value = "B" : ws.Range("C4").Value = 420 : ws.Range("D4").Value = 9
    ws.Range("A5").Value = "東區" : ws.Range("B5").Value = "A" : ws.Range("C5").Value = 610 : ws.Range("D5").Value = 12
    ws.Range("A6").Value = "南區" : ws.Range("B6").Value = "A" : ws.Range("C6").Value = 280 : ws.Range("D6").Value = 5
    ws.Range("A7").Value = "東區" : ws.Range("B7").Value = "B" : ws.Range("C7").Value = 390 : ws.Range("D7").Value = 8

    ws.Columns("A:D").AutoFit
End Sub
