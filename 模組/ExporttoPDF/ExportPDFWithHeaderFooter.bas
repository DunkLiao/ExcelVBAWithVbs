Attribute VB_Name = "ExportPDFWithHeaderFooter"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithHeaderFooter
'功能說明: 為工作表設定自訂頁首頁尾後再匯出為 PDF 的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestExportPDFWithHeaderFooter()
    Call ExportPDFWithHeaderFooter
End Sub

' 設定頁首頁尾後匯出 PDF
Sub ExportPDFWithHeaderFooter()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim pdfPath As String

    Set ws = ActiveSheet

    ' 套用頁首頁尾設定
    With ws.PageSetup
        .LeftHeader = "機密文件"
        .CenterHeader = ws.Name
        .RightHeader = Format(Now(), "yyyy/mm/dd")
        .LeftFooter = "版本：1.0"
        .CenterFooter = "第 &P 頁 / 共 &N 頁"
        .RightFooter = "Dunk"
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
    End With

    ' 設定輸出路徑
    pdfPath = ThisWorkbook.Path & "\" & ws.Name & "_含頁首頁尾.pdf"

    ' 匯出 PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "PDF 已匯出至：" & vbCrLf & pdfPath, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "匯出 PDF 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
