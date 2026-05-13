Attribute VB_Name = "ExportPDFByNamedRange"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFByNamedRange
'功能說明: 依活頁簿中定義的具名範圍，將每個具名範圍匯出為獨立 PDF 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub ExportAllNamedRangesToPDF()
    On Error GoTo ErrHandler
    Dim wb          As Workbook
    Dim nm          As Name
    Dim rng         As Range
    Dim outputDir   As String
    Dim filePath    As String
    Dim safeName    As String
    Dim exportCount As Long

    Set wb = ThisWorkbook
    outputDir = wb.Path & "\NamedRangePDF\"
    If Dir(outputDir, vbDirectory) = "" Then MkDir outputDir
    exportCount = 0

    For Each nm In wb.Names
        On Error Resume Next
        Set rng = nm.RefersToRange
        On Error GoTo ErrHandler
        If Not rng Is Nothing Then
            safeName = nm.Name
            If InStr(safeName, "!") > 0 Then
                safeName = Mid(safeName, InStr(safeName, "!") + 1)
            End If
            safeName = Replace(safeName, "/", "-")
            safeName = Replace(safeName, ":", "-")
            safeName = Replace(safeName, "*", "-")
            safeName = Replace(safeName, "?", "-")
            filePath = outputDir & safeName & ".pdf"
            rng.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=filePath, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
            exportCount = exportCount + 1
            Set rng = Nothing
        End If
    Next nm
    If exportCount = 0 Then
        MsgBox "活頁簿中未找到有效的具名範圍。", vbExclamation, "提示"
    Else
        MsgBox "已成功匯出 " & exportCount & " 個具名範圍為 PDF，儲存於：" & outputDir, vbInformation, "完成"
    End If
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Public Sub CreateSampleNamedRanges()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("具名範圍測試")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "具名範圍測試"
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "銷售額"
    ws.Range("A2").Value = "一月" : ws.Range("B2").Value = 120000
    ws.Range("A3").Value = "二月" : ws.Range("B3").Value = 98000
    ws.Range("A4").Value = "三月" : ws.Range("B4").Value = 135000
    ThisWorkbook.Names.Add Name:="銷售報表範圍", RefersTo:=ws.Range("A1:B4")
    ws.Range("D1").Value = "項目"
    ws.Range("E1").Value = "金額"
    ws.Range("D2").Value = "差旅費" : ws.Range("E2").Value = 15000
    ws.Range("D3").Value = "辦公費" : ws.Range("E3").Value = 8000
    ws.Range("D4").Value = "雜支"   : ws.Range("E4").Value = 3000
    ThisWorkbook.Names.Add Name:="費用報表範圍", RefersTo:=ws.Range("D1:E4")
    ws.Columns.AutoFit
    MsgBox "已建立兩個具名範圍，可執行 ExportAllNamedRangesToPDF 進行匯出。", vbInformation, "完成"
End Sub

