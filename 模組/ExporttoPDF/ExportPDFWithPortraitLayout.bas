Attribute VB_Name = "ExportPDFWithPortraitLayout"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithPortraitLayout
'功能說明: 以縱向版面（Portrait）將現用工作表匯出為 PDF 並指定儲存路徑
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestExportPDFWithPortraitLayout()
    Call ExportPDFWithPortraitLayout
End Sub

' 以縱向版面匯出 PDF
Sub ExportPDFWithPortraitLayout()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim savePath As String
    Dim fd As FileDialog

    Set ws = ActiveSheet

    ' 讓使用者選擇儲存位置
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    fd.Title = "請選擇 PDF 儲存位置"
    fd.InitialFileName = ws.Name & "_縱向.pdf"
    fd.FilterIndex = 1

    If fd.Show <> -1 Then
        MsgBox "未選擇儲存位置，程式結束。", vbInformation, "取消"
        Exit Sub
    End If
    savePath = fd.SelectedItems(1)
    If LCase(Right(savePath, 4)) <> ".pdf" Then
        savePath = savePath & ".pdf"
    End If

    ' 設定縱向版面
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
    End With

    ' 匯出 PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "PDF 已匯出至：" & vbCrLf & savePath, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "匯出 PDF 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
