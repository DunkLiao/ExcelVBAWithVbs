Attribute VB_Name = "PivotBoxWhiskerChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotBoxWhiskerChartExample
'功能說明: 以樞紐分析表為基礎建立盒鬚圖（Box and Whisker）樞紐分析圖範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestPivotBoxWhiskerChart()
    Call CreatePivotBoxWhiskerChart
End Sub

' 建立盒鬚圖樞紐分析圖
Sub CreatePivotBoxWhiskerChart()
    On Error GoTo ErrorHandler

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim lastRow As Long
    Dim dataRange As Range

    Set wsData = GetOrCreateSheet(ThisWorkbook, "盒鬚圖資料")
    Call FillBoxWhiskerData(wsData)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "盒鬚圖樞紐")

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1:B" & lastRow))

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="盒鬚圖樞紐")

    With pt.PivotFields("部門")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields("分數")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "分數合計"
    End With

    pt.TableStyle2 = "PivotStyleMedium4"

    Set chartObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("E3").Left, _
        Top:=wsPivot.Range("E3").Top, _
        Width:=400, _
        Height:=300)

    Set cht = chartObj.Chart
    Set dataRange = wsData.Range("A1:B" & lastRow)
    cht.SetSourceData Source:=dataRange

    ' 嘗試設定盒鬚圖類型（Excel 2016+，xlBoxwhisker=120）
    On Error Resume Next
    cht.ChartType = 120
    If Err.Number <> 0 Then
        cht.ChartType = xlColumnClustered
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    cht.HasTitle = True
    cht.ChartTitle.Text = "各部門分數分佈（盒鬚圖）"
    cht.HasLegend = False

    wsPivot.Columns.AutoFit

    MsgBox "盒鬚圖樞紐分析圖已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立盒鬚圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 填入盒鬚圖範例資料
Private Sub FillBoxWhiskerData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "分數"
    ws.Range("A1:B1").Font.Bold = True

    Dim departments As Variant
    Dim scores As Variant
    Dim i As Integer

    departments = Array("業務", "工程", "行銷", "業務", "工程", "行銷", _
                        "業務", "工程", "行銷", "業務", "工程", "行銷", _
                        "業務", "工程", "行銷")
    scores = Array(78, 85, 72, 90, 68, 88, 65, 92, 75, 82, 79, 95, 70, 84, 69)

    For i = 0 To UBound(departments)
        ws.Cells(i + 2, 1).Value = departments(i)
        ws.Cells(i + 2, 2).Value = scores(i)
    Next i

    ws.Columns("A:B").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
