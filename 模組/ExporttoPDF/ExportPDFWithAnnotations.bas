Option Explicit
Attribute VB_Name = "ExportPDFWithAnnotations"
'*************************************************************************************
'模組名稱: 含附註匯出 PDF
'功能說明: 將使用中工作表（含儲存格批注）匯出為 PDF，並在頁尾加上匯出時間
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub ExportPDFWithAnnotations()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim pdfPath As String
    Dim savePath As String

    Set ws = ActiveSheet

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=ws.Name & "_含附註.pdf", _
        FileFilter:="PDF 檔案 (*.pdf),*.pdf")

    If savePath = "False" Or savePath = "" Then Exit Sub

    ' 備份頁首頁尾設定
    Dim origFooter As String
    origFooter = ws.PageSetup.CenterFooter

    ' 設定頁尾顯示匯出時間
    ws.PageSetup.CenterFooter = "匯出時間：" & Format(Now, "yyyy/mm/dd hh:mm:ss")

    ' 設定顯示批注（列印時顯示）
    ws.PageSetup.PrintComments = xlPrintSheetEnd

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' 還原頁尾與批注設定
    ws.PageSetup.CenterFooter = origFooter
    ws.PageSetup.PrintComments = xlPrintNoComments

    MsgBox "含附註 PDF 已匯出至：" & savePath, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    ws.PageSetup.PrintComments = xlPrintNoComments
    MsgBox "匯出 PDF 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
