Option Explicit
Attribute VB_Name = "ExportPDFWithLandscapeLayout"
'*************************************************************************************
'模組名稱: ExportPDFWithLandscapeLayout
'功能說明: 以橫向版面配置將目前工作表匯出為 PDF 檔案，匯出後還原版面設定
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub ExportPDFWithLandscapeLayout()
    Dim ws As Worksheet
    Dim savePath As String
    Dim originalOrientation As XlPageOrientation
    Dim originalFitWide As Variant
    Dim originalFitTall As Variant

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.ActiveSheet

    originalOrientation = ws.PageSetup.Orientation
    originalFitWide = ws.PageSetup.FitToPagesWide
    originalFitTall = ws.PageSetup.FitToPagesTall

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
    End With

    If ThisWorkbook.Path <> "" Then
        savePath = ThisWorkbook.Path & "\" & ws.Name & "_橫向版面.pdf"
    Else
        savePath = Environ("USERPROFILE") & "\Desktop\" & ws.Name & "_橫向版面.pdf"
    End If

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ws.PageSetup.Orientation = originalOrientation
    ws.PageSetup.FitToPagesWide = originalFitWide
    ws.PageSetup.FitToPagesTall = originalFitTall

    MsgBox "已匯出橫向 PDF：" & vbCrLf & savePath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "匯出橫向 PDF 失敗"
End Sub