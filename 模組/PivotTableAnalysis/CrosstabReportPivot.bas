Attribute VB_Name = "CrosstabReportPivot"
Option Explicit
'*************************************************************************************
'模組名稱: CrosstabReportPivot
'功能說明: 建立雙維度交叉報表樞紐分析表，同時呈現列與欄的分類統計
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestCrosstabReportPivot()
    Call CreateCrosstabReportPivot
End Sub

' 建立交叉報表樞紐分析表
Sub CreateCrosstabReportPivot()
    On Error GoTo ErrorHandler

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim lastRow As Long

    Set wsData = GetOrCreateSheet(ThisWorkbook, "交叉報表資料")
    Call FillCrosstabData(wsData)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "交叉樞紐分析")

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1:C" & lastRow))

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="交叉報表")

    With pt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields("產品類別")
        .Orientation = xlColumnField
        .Position = 1
    End With

    With pt.PivotFields("銷售額")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0"
        .Name = "銷售額合計"
    End With

    pt.TableStyle2 = "PivotStyleMedium9"
    wsPivot.Columns.AutoFit

    MsgBox "交叉報表樞紐分析表已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立交叉報表時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 填入交叉報表範例資料
Private Sub FillCrosstabData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "產品類別"
    ws.Range("C1").Value = "銷售額"
    ws.Range("A1:C1").Font.Bold = True

    Dim regions As Variant
    Dim categories As Variant
    Dim r As Integer
    Dim c As Integer
    Dim rowNum As Integer

    regions = Array("北部", "中部", "南部")
    categories = Array("電子", "服飾", "食品", "家具")

    rowNum = 2
    For r = 0 To UBound(regions)
        For c = 0 To UBound(categories)
            ws.Cells(rowNum, 1).Value = regions(r)
            ws.Cells(rowNum, 2).Value = categories(c)
            ws.Cells(rowNum, 3).Value = (r + 1) * (c + 1) * 5000 + 10000
            rowNum = rowNum + 1
        Next c
    Next r

    ws.Columns("A:C").AutoFit
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
