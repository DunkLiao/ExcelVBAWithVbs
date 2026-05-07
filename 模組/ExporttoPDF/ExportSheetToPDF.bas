Attribute VB_Name = "ExportSheetToPDF"
Option Explicit
'*************************************************************************************
'模組名稱: ExportSheetToPDF
'功能說明: 將指定工作表或整個活頁簿匯出為 PDF 檔案
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口
Sub TestExportSheetToPDF()
    Call ExportActiveSheetToPDF
End Sub

' 將目前使用中的工作表匯出為 PDF
Sub ExportActiveSheetToPDF()
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim defaultName As String

    Set ws = ActiveSheet
    defaultName = ws.Name & ".pdf"

    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請選擇 PDF 儲存路徑"
        .InitialFileName = defaultName
        .FilterIndex = 1
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        pdfPath = .SelectedItems(1)
    End With

    If LCase(Right(pdfPath, 4)) <> ".pdf" Then
        pdfPath = pdfPath & ".pdf"
    End If

    On Error GoTo ErrorHandler
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "PDF 匯出完成！" & vbCrLf & "儲存路徑：" & pdfPath, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "匯出失敗！" & vbCrLf & "錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 將整個活頁簿匯出為單一 PDF
Sub ExportWorkbookToPDF()
    Dim pdfPath As String
    Dim defaultName As String

    defaultName = ThisWorkbook.Name
    If InStr(defaultName, ".") > 0 Then
        defaultName = Left(defaultName, InStrRev(defaultName, ".") - 1) & ".pdf"
    End If

    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請選擇整本活頁簿 PDF 儲存路徑"
        .InitialFileName = defaultName
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        pdfPath = .SelectedItems(1)
    End With

    If LCase(Right(pdfPath, 4)) <> ".pdf" Then
        pdfPath = pdfPath & ".pdf"
    End If

    On Error GoTo ErrorHandler
    ThisWorkbook.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "整本活頁簿 PDF 匯出完成！" & vbCrLf & "儲存路徑：" & pdfPath, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "匯出失敗！" & vbCrLf & "錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次將所有工作表分別匯出為獨立 PDF
Sub ExportEachSheetToPDF()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim pdfPath As String
    Dim count As Integer

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    count = 0

    For Each ws In ThisWorkbook.Worksheets
        pdfPath = folderPath & "" & ws.Name & ".pdf"
        On Error Resume Next
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=pdfPath, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        If Err.Number = 0 Then count = count + 1
        On Error GoTo 0
    Next ws

    Application.ScreenUpdating = True
    MsgBox "批次匯出完成！共匯出 " & count & " 個 PDF 檔案。", vbInformation, "完成"
End Sub
