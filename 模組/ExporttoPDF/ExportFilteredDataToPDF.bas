Attribute VB_Name = "ExportFilteredDataToPDF"
Option Explicit
'*************************************************************************************
'模組名稱: ExportFilteredDataToPDF
'功能說明: 將工作表中篩選後的可見資料列複製到暫存工作表後匯出為 PDF
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestExportFilteredDataToPDF()
    Dim ws As Worksheet
    Set ws = GetOrCreateFilterPdfWs(ThisWorkbook, "篩選PDF測試")
    ws.Cells.Clear

    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "業績"
    ws.Range("A2").Value = "王小明" : ws.Range("B2").Value = "業務部" : ws.Range("C2").Value = 85000
    ws.Range("A3").Value = "李大華" : ws.Range("B3").Value = "行銷部" : ws.Range("C3").Value = 72000
    ws.Range("A4").Value = "陳美玲" : ws.Range("B4").Value = "業務部" : ws.Range("C4").Value = 91000
    ws.Range("A5").Value = "林俊傑" : ws.Range("B5").Value = "研發部" : ws.Range("C5").Value = 68000

    ' 套用自動篩選：僅顯示業務部
    ws.Range("A1:C5").AutoFilter Field:=2, Criteria1:="業務部"

    Call ExportFilteredDataToPDF(ws)
End Sub

' 將篩選後可見資料匯出為 PDF
Sub ExportFilteredDataToPDF(ByVal wsSource As Worksheet)
    On Error GoTo ErrorHandler

    Dim wsTmp As Worksheet
    Dim pdfPath As String
    Dim visibleRows As Long

    ' 計算可見列數（含標題）
    Dim cell As Range
    visibleRows = 0
    For Each cell In wsSource.UsedRange.Rows
        If cell.EntireRow.Hidden = False Then
            visibleRows = visibleRows + 1
        End If
    Next cell

    If visibleRows <= 1 Then
        MsgBox "篩選後無資料可匯出。", vbInformation, "提示"
        Exit Sub
    End If

    ' 建立暫存工作表並貼上可見資料
    Set wsTmp = ThisWorkbook.Worksheets.Add
    wsTmp.Name = "PDF暫存_" & Format(Now, "hhmmss")

    Application.ScreenUpdating = False

    wsSource.UsedRange.SpecialCells(xlCellTypeVisible).Copy
    wsTmp.Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    Application.CutCopyMode = False
    wsTmp.UsedRange.Columns.AutoFit

    ' 設定 PDF 路徑
    pdfPath = ThisWorkbook.Path & "\篩選資料_" & _
        wsSource.Name & "_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"

    ' 匯出 PDF
    wsTmp.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False

    ' 刪除暫存工作表
    Application.DisplayAlerts = False
    wsTmp.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "篩選資料已匯出 PDF：" & vbCrLf & pdfPath, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wsTmp Is Nothing Then wsTmp.Delete
    On Error GoTo 0
    MsgBox "匯出 PDF 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateFilterPdfWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateFilterPdfWs = ws
End Function
