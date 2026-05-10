Attribute VB_Name = "CompactFormPivot"
Option Explicit
'*************************************************************************************
'模組名稱: CompactFormPivot
'功能說明: 建立樞紐分析表並示範精簡、大綱、表格三種版面配置的切換
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestCompactFormPivot()
    Call CreateCompactFormPivotExample
End Sub

' 建立樞紐分析表並示範版面配置切換
Sub CreateCompactFormPivotExample()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim dataWs As Worksheet
    Dim pivotWs As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range

    Set wb = ThisWorkbook
    Set dataWs = GetOrCreateSheet(wb, "樞紐版面資料")
    Set pivotWs = GetOrCreateSheet(wb, "樞紐版面配置範例")

    Call FillPivotData(dataWs)
    Set dataRange = dataWs.Range("A1:D" & dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row)

    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="版面配置樞紐")

    With pt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields("類別")
        .Orientation = xlRowField
        .Position = 2
    End With

    With pt.PivotFields("季度")
        .Orientation = xlColumnField
        .Position = 1
    End With

    With pt.PivotFields("銷售額")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0"
        .Name = "加總-銷售額"
    End With

    pt.RowAxisLayout xlCompactRow

    pt.ColumnGrand = True
    pt.RowGrand = True

    pivotWs.Range("A1").Value = "目前版面：精簡配置（xlCompactRow）"
    pivotWs.Range("A1").Font.Bold = True
    pivotWs.Range("A2").Value = "可透過 pt.RowAxisLayout xlOutlineRow 或 xlTabularRow 切換版面"

    pivotWs.Activate
    MsgBox "樞紐分析表建立完成！" & vbCrLf & _
           "目前為「精簡配置」。" & vbCrLf & _
           "可在程式中改用 xlOutlineRow（大綱）或 xlTabularRow（表格）。", _
           vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立樞紐分析表時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 切換為大綱版面配置
Sub SwitchToOutlineLayout()
    On Error GoTo ErrorHandler
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    pt.RowAxisLayout xlOutlineRow
    MsgBox "已切換為「大綱配置」。", vbInformation, "完成"
    Exit Sub
ErrorHandler:
    MsgBox "切換版面配置時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 切換為表格版面配置
Sub SwitchToTabularLayout()
    On Error GoTo ErrorHandler
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    pt.RowAxisLayout xlTabularRow
    MsgBox "已切換為「表格配置」。", vbInformation, "完成"
    Exit Sub
ErrorHandler:
    MsgBox "切換版面配置時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 填入樞紐範例資料
Private Sub FillPivotData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "類別"
    ws.Range("C1").Value = "季度"
    ws.Range("D1").Value = "銷售額"
    ws.Range("A1:D1").Font.Bold = True

    Dim data As Variant
    data = Array( _
        Array("北部", "電子", "Q1", 85000), Array("北部", "電子", "Q2", 92000), _
        Array("北部", "食品", "Q1", 45000), Array("北部", "食品", "Q2", 48000), _
        Array("中部", "電子", "Q1", 62000), Array("中部", "電子", "Q2", 70000), _
        Array("中部", "食品", "Q1", 38000), Array("中部", "食品", "Q2", 41000), _
        Array("南部", "電子", "Q1", 55000), Array("南部", "電子", "Q2", 61000), _
        Array("南部", "食品", "Q1", 32000), Array("南部", "食品", "Q2", 35000) _
    )

    Dim i As Integer
    For i = 0 To 11
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
        ws.Cells(i + 2, 4).Value = data(i)(3)
    Next i

    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
