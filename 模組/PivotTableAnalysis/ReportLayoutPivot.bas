Attribute VB_Name = "ReportLayoutPivot"
Option Explicit
'*************************************************************************************
'模組名稱: ReportLayoutPivot
'功能說明: 示範如何以 VBA 切換樞紐分析表的版面配置（壓縮/大綱/表格式），
'          並設定重複項目標籤、小計位置與空白列等報表版面選項
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

' 主程式：建立範例樞紐並展示三種版面
Sub DemoPivotReportLayout()
    Dim ws      As Worksheet
    Dim wsPivot As Worksheet
    Dim pt      As PivotTable
    Dim pc      As PivotCache

    ' 建立範例資料
    Set ws = GetOrCreateLayoutSheet("版面資料來源")
    Call FillLayoutSampleData(ws)

    ' 移除舊樞紐工作表
    Dim wsOld As Worksheet
    On Error Resume Next
    Set wsOld = ThisWorkbook.Worksheets("版面示範")
    On Error GoTo 0
    If Not wsOld Is Nothing Then
        Application.DisplayAlerts = False
        wsOld.Delete
        Application.DisplayAlerts = True
    End If

    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "版面示範"

    ' 建立 PivotCache
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)))

    ' 建立樞紐（壓縮版面）
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("B2"), _
        TableName:="版面示範樞紐")

    With pt
        .PivotFields("地區").Orientation = xlRowField
        .PivotFields("地區").Position = 1
        .PivotFields("業務員").Orientation = xlRowField
        .PivotFields("業務員").Position = 2
        .PivotFields("銷售金額").Orientation = xlDataField
        .PivotFields("銷售金額").Function = xlSum
        .PivotFields("銷售金額").NumberFormat = "#,##0"
        .TableStyle2 = "PivotStyleMedium7"
    End With

    ' 預設為壓縮版面
    Call SetLayoutCompact(pt)

    wsPivot.Columns.AutoFit

    ' 詢問使用者選擇版面
    Dim choice As Integer
    choice = MsgBox("樞紐分析表已建立（目前：壓縮版面）。" & vbCrLf & _
                    "是否切換為表格式版面？", vbYesNo + vbQuestion, "版面選擇")
    If choice = vbYes Then
        Call SetLayoutTabular(pt)
        wsPivot.Columns.AutoFit
        MsgBox "已切換為表格式版面！", vbInformation, "完成"
    Else
        MsgBox "保持壓縮版面。", vbInformation, "完成"
    End If
End Sub

' 設定壓縮版面（Compact Form）
Sub SetLayoutCompact(ByVal pt As PivotTable)
    pt.RowAxisLayout xlCompactRow
    Dim pf As PivotField
    For Each pf In pt.RowFields
        pf.RepeatLabels = False
    Next pf
End Sub

' 設定大綱版面（Outline Form）
Sub SetLayoutOutline(ByVal pt As PivotTable)
    pt.RowAxisLayout xlOutlineRow
    Dim pf As PivotField
    For Each pf In pt.RowFields
        pf.RepeatLabels = False
        pf.ShowDetail = True
    Next pf
End Sub

' 設定表格式版面（Tabular Form），重複標籤並移除小計
Sub SetLayoutTabular(ByVal pt As PivotTable)
    pt.RowAxisLayout xlTabularRow
    Dim pf As PivotField
    For Each pf In pt.RowFields
        pf.RepeatLabels = True
        pf.Subtotals = Array(False, False, False, False, False, False, _
                             False, False, False, False, False, False)
    Next pf
End Sub

Private Sub FillLayoutSampleData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("地區", "業務員", "銷售金額")
    ws.Range("A2:C2").Value = Array("北區", "張志豪", 85000)
    ws.Range("A3:C3").Value = Array("北區", "李佳蓉", 92000)
    ws.Range("A4:C4").Value = Array("中區", "王大明", 73000)
    ws.Range("A5:C5").Value = Array("中區", "陳雅婷", 68000)
    ws.Range("A6:C6").Value = Array("南區", "林志偉", 95000)
    ws.Range("A7:C7").Value = Array("南區", "黃淑芬", 80000)
    ws.Range("A8:C8").Value = Array("北區", "張志豪", 60000)
    ws.Range("A9:C9").Value = Array("中區", "王大明", 55000)
    ws.Range("A10:C10").Value = Array("南區", "林志偉", 77000)
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateLayoutSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateLayoutSheet = ws
End Function
