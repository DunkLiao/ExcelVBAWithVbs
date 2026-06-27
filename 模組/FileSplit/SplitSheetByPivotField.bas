Attribute VB_Name = "SplitSheetByPivotField"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetByPivotField
'功能說明: 依樞紐分析表欄位值分割資料至個別工作表的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestSplitByPivotField()
    Call SplitSheetByPivotFieldDemo
End Sub

Sub SplitSheetByPivotFieldDemo()
    Dim wsSource As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem

    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsSource = ThisWorkbook.Worksheets("分割來源資料")
    If Not wsSource Is Nothing Then wsSource.Delete
    Set wsPivot = ThisWorkbook.Worksheets("樞紐分割暫存")
    If Not wsPivot Is Nothing Then wsPivot.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 建立來源資料
    Set wsSource = ThisWorkbook.Worksheets.Add
    wsSource.Name = "分割來源資料"

    wsSource.Range("A1").Value = "部門"
    wsSource.Range("B1").Value = "員工姓名"
    wsSource.Range("C1").Value = "業績金額"
    wsSource.Range("A1:C1").Font.Bold = True

    wsSource.Range("A2").Value = "業務部"
    wsSource.Range("B2").Value = "張三"
    wsSource.Range("C2").Value = 85000

    wsSource.Range("A3").Value = "業務部"
    wsSource.Range("B3").Value = "李四"
    wsSource.Range("C3").Value = 72000

    wsSource.Range("A4").Value = "工程部"
    wsSource.Range("B4").Value = "王五"
    wsSource.Range("C4").Value = 68000

    wsSource.Range("A5").Value = "工程部"
    wsSource.Range("B5").Value = "趙六"
    wsSource.Range("C5").Value = 91000

    wsSource.Range("A6").Value = "財務部"
    wsSource.Range("B6").Value = "孫七"
    wsSource.Range("C6").Value = 77000

    wsSource.Columns("A:C").AutoFit

    ' 建立樞紐分析表
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "樞紐分割暫存"

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsSource.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="分割用樞紐")

    pt.AddFields RowFields:="部門"

    With pt.PivotFields("業績金額")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "合計業績"
    End With

    ' 使用篩選欄位進行分割
    Set pf = pt.PageFields.Add(pt.PivotFields("部門"))

    ' 遍歷各部門並產生對應的工作表
    Dim newWs As Worksheet
    For Each pi In pf.PivotItems
        pf.CurrentPage = pi.Value

        Set newWs = ThisWorkbook.Worksheets.Add
        newWs.Name = pi.Value

        With wsSource.Range("A1").CurrentRegion
            .AutoFilter Field:=1, Criteria1:=pi.Value
            .SpecialCells(xlCellTypeVisible).Copy newWs.Range("A1")
            .AutoFilter
        End With

        newWs.Columns("A:C").AutoFit
    Next pi

    ' 隱藏暫存樞紐工作表
    wsPivot.Visible = xlSheetHidden
    wsSource.Visible = xlSheetHidden

    MsgBox "已依部門分割完成，請查看各個新工作表。", vbInformation, "完成"
End Sub
