Attribute VB_Name = "ExportPDFWithPageNumber"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithPageNumber
'功能說明: 設定頁尾顯示頁碼後，將工作表匯出為PDF
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub ExportPDFWithPageNumber()
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim savePath As String

    Set ws = ActiveSheet

    With ws.PageSetup
        .CenterFooter = "第 &P 頁，共 &N 頁"
        .RightFooter = "匯出日期：" & Format(Now(), "yyyy/mm/dd")
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    savePath = ThisWorkbook.Path
    If savePath = "" Then savePath = Environ("USERPROFILE") & "\Desktop"
    pdfPath = savePath & "\" & ws.Name & "_含頁碼_" & Format(Now(), "yyyymmdd_HHmmss") & ".pdf"

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "PDF已匯出至：" & vbCrLf & pdfPath, vbInformation, "完成"
End Sub
