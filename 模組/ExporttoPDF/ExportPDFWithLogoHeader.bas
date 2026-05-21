Option Explicit
Attribute VB_Name = "ExportPDFWithLogoHeader"
'*************************************************************************************
'模組名稱: ExportPDFWithLogoHeader
'功能說明: 匯出 PDF 前，在工作表頁首區域插入公司標誌圖片，再以固定格式輸出
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestExportPDFWithLogoHeader()
    Dim logoPath As String
    logoPath = Application.GetOpenFilename( _
        "圖片檔案 (*.png;*.jpg;*.bmp),*.png;*.jpg;*.bmp", _
        1, "請選擇公司標誌圖片", , False)

    If VarType(logoPath) = vbBoolean Then
        MsgBox "已取消操作。", vbInformation, "提示"
        Exit Sub
    End If

    Dim outputPath As String
    outputPath = Application.GetSaveAsFilename( _
        "輸出報表.pdf", "PDF 檔案 (*.pdf),*.pdf", 1, "請指定 PDF 輸出路徑")

    If VarType(outputPath) = vbBoolean Then
        MsgBox "已取消操作。", vbInformation, "提示"
        Exit Sub
    End If

    Call ExportWithLogoHeader(ActiveSheet, logoPath, outputPath)
End Sub

Sub ExportWithLogoHeader( _
    ByVal ws As Worksheet, _
    ByVal logoPath As String, _
    ByVal outputPath As String)

    On Error GoTo ErrorHandler

    Dim imgObj As Picture
    Set imgObj = ws.Pictures.Insert(logoPath)

    With imgObj
        .Left = ws.Range("A1").Left
        .Top = ws.Range("A1").Top
        .Width = 120
        .Height = 40
        .Placement = xlMoveAndSize
    End With

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    imgObj.Delete

    MsgBox "含標誌頁首的 PDF 已匯出至：" & vbCrLf & outputPath, _
        vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "匯出含標誌 PDF 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
