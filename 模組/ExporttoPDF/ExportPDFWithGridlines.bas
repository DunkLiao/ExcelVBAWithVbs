Attribute VB_Name = "ExportPDFWithGridlines"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithGridlines
'功能說明: 將作用中工作表以「顯示格線」的設定匯出為 PDF，
'          匯出後自動還原原始格線設定，避免更動使用者設定
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

Sub ExportSheetPDFWithGridlines()
    Dim ws              As Worksheet
    Dim savePath        As String
    Dim origGridlines   As Boolean
    Dim origPrintGrid   As Boolean

    Set ws = ActiveSheet

    ' 選擇儲存路徑
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "另存 PDF（含格線）"
        .InitialFileName = ws.Name & "_含格線.pdf"
        .FilterIndex = 2
        If .Show <> -1 Then Exit Sub
        savePath = .SelectedItems(1)
    End With

    ' 確保副檔名
    If LCase(Right(savePath, 4)) <> ".pdf" Then
        savePath = savePath & ".pdf"
    End If

    ' 記錄原始設定
    origGridlines = ws.DisplayGridlines
    origPrintGrid = ws.PageSetup.PrintGridlines

    ' 強制顯示並列印格線
    ws.DisplayGridlines = True
    ws.PageSetup.PrintGridlines = True

    ' 匯出 PDF
    On Error GoTo ErrHandler
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        FileName:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' 還原設定
    ws.DisplayGridlines = origGridlines
    ws.PageSetup.PrintGridlines = origPrintGrid

    MsgBox "PDF 已匯出（含格線）：" & vbCrLf & savePath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    ws.DisplayGridlines = origGridlines
    ws.PageSetup.PrintGridlines = origPrintGrid
    MsgBox "匯出失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次匯出活頁簿所有工作表（含格線）
Sub ExportAllSheetsPDFWithGridlines()
    Dim ws          As Worksheet
    Dim savePath    As String
    Dim outputDir   As String
    Dim origGrid    As Boolean
    Dim origPrint   As Boolean

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then Exit Sub
        outputDir = .SelectedItems(1)
    End With

    If Right(outputDir, 1) <> "" Then outputDir = outputDir & ""

    Dim successCount As Long
    successCount = 0

    For Each ws In ThisWorkbook.Worksheets
        origGrid = ws.DisplayGridlines
        origPrint = ws.PageSetup.PrintGridlines
        ws.DisplayGridlines = True
        ws.PageSetup.PrintGridlines = True

        savePath = outputDir & ws.Name & "_含格線.pdf"
        On Error Resume Next
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            FileName:=savePath, _
            Quality:=xlQualityStandard, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        On Error GoTo 0

        ws.DisplayGridlines = origGrid
        ws.PageSetup.PrintGridlines = origPrint
        successCount = successCount + 1
    Next ws

    MsgBox "共匯出 " & successCount & " 個工作表 PDF（含格線）至：" & outputDir, _
           vbInformation, "完成"
End Sub
