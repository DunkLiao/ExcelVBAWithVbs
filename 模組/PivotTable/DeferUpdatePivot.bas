Attribute VB_Name = "DeferUpdatePivot"
Option Explicit
'*************************************************************************************
'模組名稱: DeferUpdatePivot
'功能說明: 延遲版面配置更新模式示範（ManualUpdate）
'          在大量欄位操作時，先關閉自動更新以提升效能
'          完成所有欄位設定後再統一刷新樞紐分析表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestDeferUpdatePivot()
    Call CreateDeferUpdatePivot
End Sub

' 建立樞紐分析表並示範 ManualUpdate 延遲更新模式
Sub CreateDeferUpdatePivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim startTime As Double
    Dim endTime As Double

    Set wsData = GetOrCreateSheet(ThisWorkbook, "大量欄位資料")
    Call FillLargeFieldData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "延遲更新樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="延遲更新樞紐分析表")

    startTime = Timer

    ' ── 開啟延遲更新模式（ManualUpdate = True）──
    ' 此後所有欄位設定均不觸發重算，直到 ManualUpdate = False
    pt.ManualUpdate = True

    ' 批次設定多個欄位（不觸發重算）
    With pt.PivotFields("部門")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields("月份")
        .Orientation = xlColumnField
        .Position = 1
    End With

    pt.AddDataField pt.PivotFields("薪資"), "薪資合計", xlSum
    pt.AddDataField pt.PivotFields("獎金"), "獎金合計", xlSum

    pt.RowAxisLayout xlTabularRow
    pt.TableStyle2 = "PivotStyleMedium4"

    ' ── 關閉延遲更新，觸發一次性重算 ──
    pt.ManualUpdate = False

    endTime = Timer

    wsPivot.Columns("A:I").AutoFit
    wsPivot.Activate
    MsgBox "延遲更新模式示範完成！" & Chr(13) & _
           "所有欄位設定完畢後才進行一次重算。" & Chr(13) & _
           "耗時：" & Format(endTime - startTime, "0.000") & " 秒", _
           vbInformation, "完成"
End Sub

' 填入大量欄位測試資料
Private Sub FillLargeFieldData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "月份"
    ws.Range("C1").Value = "薪資"
    ws.Range("D1").Value = "獎金"

    Dim depts As Variant
    Dim months As Variant
    depts = Array("業務部", "研發部", "管理部", "財務部")
    months = Array("一月", "二月", "三月", "四月", "五月", "六月")

    Dim row As Integer
    Dim d As Integer
    Dim m As Integer
    row = 2
    For d = 0 To UBound(depts)
        For m = 0 To UBound(months)
            ws.Cells(row, 1).Value = depts(d)
            ws.Cells(row, 2).Value = months(m)
            ws.Cells(row, 3).Value = 30000 + (d * 5000) + (m * 1000)
            ws.Cells(row, 4).Value = 5000 + (d * 1000) + (m * 500)
            row = row + 1
        Next m
    Next d

    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表，並清除內容
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
