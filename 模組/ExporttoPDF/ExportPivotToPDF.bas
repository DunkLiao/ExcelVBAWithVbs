Option Explicit
Attribute VB_Name = "ExportPivotToPDF"
'*************************************************************************************
'模組名稱: ExportPivotToPDF
'功能說明: 將工作表中的樞紐分析表單獨匯出為 PDF 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestExportPivotToPDF()
    Call ExportPivotTableToPDF
End Sub

' 匯出作用工作表的樞紐分析表為 PDF
Sub ExportPivotTableToPDF()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim ptRange As Range
    Dim savePath As String
    Dim wsTemp As Worksheet
    Dim defaultPath As String

    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "請先選取一個工作表。", vbExclamation, "提示"
        Exit Sub
    End If

    ' 檢查是否存在樞紐分析表
    If ws.PivotTables.Count = 0 Then
        ' 若無樞紐分析表，建立示範用樞紐
        Set ws = GetOrCreateWorksheet("樞紐PDF範例")
        ws.Cells.Clear
        ws.Range("A1").Value = "部門"
        ws.Range("B1").Value = "金額"
        ws.Range("A2").Value = "業務部"
        ws.Range("B2").Value = 500
        ws.Range("A3").Value = "行銷部"
        ws.Range("B3").Value = 300
        ws.Range("A4").Value = "研發部"
        ws.Range("B4").Value = 700

        Set ptCache = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=ws.Range("A1:B4"))
        Set pt = ptCache.CreatePivotTable( _
            TableDestination:=ws.Range("D1"), _
            TableName:="示範樞紐")
        pt.PivotFields("部門").Orientation = xlRowField
        pt.PivotFields("金額").Orientation = xlDataField
    End If

    ' 取得第一個樞紐分析表的範圍
    Set pt = ws.PivotTables(1)
    Set ptRange = pt.TableRange2

    ' 選擇儲存路徑
    defaultPath = Environ("USERPROFILE") & "\Desktop\PivotExport.pdf"
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultPath, _
        FileFilter:="PDF 檔案 (*.pdf), *.pdf", _
        Title:="另存樞紐分析表為 PDF")

    If savePath = "False" Then Exit Sub

    ' 匯出為 PDF
    ptRange.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True

    MsgBox "樞紐分析表已匯出為 PDF：" & vbCrLf & savePath, vbInformation, "完成"
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
