Attribute VB_Name = "ExportPDFWithCustomPageSetup"
Option Explicit

' ============================================================
' 範例：套用自訂頁面設定後再匯出 PDF
' 功能：設定 A4 橫向、四邊 1cm 邊界、自動縮放，完成後還原
' ============================================================

Sub ExportPDFWithCustomPageSetup()
    Dim ws               As Worksheet
    Dim pdfPath          As String
    Dim origOrientation  As XlPageOrientation
    Dim origPaperSize    As XlPaperSize
    Dim origTopMargin    As Double
    Dim origBottomMargin As Double
    Dim origLeftMargin   As Double
    Dim origRightMargin  As Double
    Dim hadError         As Boolean

    Set ws = ActiveSheet

    ' 備份原始頁面設定
    With ws.PageSetup
        origOrientation  = .Orientation
        origPaperSize    = .PaperSize
        origTopMargin    = .TopMargin
        origBottomMargin = .BottomMargin
        origLeftMargin   = .LeftMargin
        origRightMargin  = .RightMargin
    End With

    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請選擇 PDF 儲存位置"
        .InitialFileName = ws.Name & "_Custom.pdf"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        pdfPath = .SelectedItems(1)
    End With

    If LCase(Right(pdfPath, 4)) <> ".pdf" Then pdfPath = pdfPath & ".pdf"

    hadError = False
    On Error GoTo RestoreSettings

    ' 套用自訂頁面設定（A4 橫向，四邊 1 公分，自動縮放至單頁寬）
    With ws.PageSetup
        .Orientation     = xlLandscape
        .PaperSize       = xlPaperA4
        .TopMargin       = Application.CentimetersToPoints(1)
        .BottomMargin    = Application.CentimetersToPoints(1)
        .LeftMargin      = Application.CentimetersToPoints(1)
        .RightMargin     = Application.CentimetersToPoints(1)
        .Zoom            = False
        .FitToPagesWide  = 1
        .FitToPagesTall  = False
    End With

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "已套用自訂頁面設定並匯出 PDF：" & vbCrLf & pdfPath, vbInformation, "完成"

RestoreSettings:
    hadError = (Err.Number <> 0)
    If hadError Then
        MsgBox "匯出失敗：" & Err.Description, vbCritical, "錯誤"
    End If
    ' 還原原始頁面設定
    On Error Resume Next
    With ws.PageSetup
        .Orientation     = origOrientation
        .PaperSize       = origPaperSize
        .TopMargin       = origTopMargin
        .BottomMargin    = origBottomMargin
        .LeftMargin      = origLeftMargin
        .RightMargin     = origRightMargin
        .Zoom            = 100
        .FitToPagesWide  = False
        .FitToPagesTall  = False
    End With
    On Error GoTo 0
End Sub
