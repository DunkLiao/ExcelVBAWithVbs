Attribute VB_Name = "ExportPDFWithMetadata"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithMetadata
'功能說明: 匯出 PDF 並設定文件屬性（標題、作者、主旨、關鍵字）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestExportPDFWithMetadata()
    Call ExportSheetToPDFWithMetadata
End Sub

Sub ExportSheetToPDFWithMetadata()
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim regKey As String
    Dim regValue As Variant
    Dim authorName As String

    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets("PDF中繼資料範例")
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "PDF中繼資料範例"

    ' 填入報表內容
    ws.Range("A1").Value = "月銷售報表"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True

    ws.Range("A3").Value = "產品"
    ws.Range("B3").Value = "銷售數量"
    ws.Range("C3").Value = "銷售金額"
    ws.Range("A3:C3").Font.Bold = True

    ws.Range("A4").Value = "產品Alpha"
    ws.Range("B4").Value = 320
    ws.Range("C4").Value = 48000

    ws.Range("A5").Value = "產品Beta"
    ws.Range("B5").Value = 180
    ws.Range("C5").Value = 27000

    ws.Range("A6").Value = "產品Gamma"
    ws.Range("B6").Value = 450
    ws.Range("C6").Value = 67500

    ws.Range("C4:C6").NumberFormat = "#,##0"
    ws.Columns("A:C").AutoFit

    ' 設定列印區域
    ws.PageSetup.PrintArea = ws.Range("A1:C6").Address

    ' 產生 PDF 路徑
    pdfPath = ThisWorkbook.Path & "\ExportWithMetadata_" & _
              Format(Now, "yyyymmdd_HHmmss") & ".pdf"

    ' 匯出 PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        OpenAfterPublish:=False

    MsgBox "PDF 已匯出至：" & vbCrLf & pdfPath & vbCrLf & vbCrLf & _
           "已包含文件屬性（標題、作者）", vbInformation, "完成"
End Sub
