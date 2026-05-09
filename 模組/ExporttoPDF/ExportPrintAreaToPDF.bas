Attribute VB_Name = "ExportPrintAreaToPDF"
Option Explicit

' ============================================================
' 範例：僅匯出工作表設定的列印範圍為 PDF
' 功能：讀取現有列印範圍設定，或由使用者選取後設定並匯出
' ============================================================

' 匯出已設定好的列印範圍為 PDF
Sub ExportPrintAreaToPDF()
    Dim ws      As Worksheet
    Dim pdfPath As String

    Set ws = ActiveSheet

    If ws.PageSetup.PrintArea = "" Then
        MsgBox "目前工作表尚未設定列印範圍。" & vbCrLf & _
               "請先選取範圍，再到「版面配置 > 列印範圍 > 設定列印範圍」。", _
               vbExclamation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請選擇列印範圍 PDF 儲存位置"
        .InitialFileName = ws.Name & "_PrintArea.pdf"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        pdfPath = .SelectedItems(1)
    End With

    If LCase(Right(pdfPath, 4)) <> ".pdf" Then pdfPath = pdfPath & ".pdf"

    On Error GoTo ErrHandler
    ' IgnorePrintAreas = False 確保只輸出已設定的列印範圍
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "列印範圍已匯出為 PDF：" & vbCrLf & pdfPath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "匯出失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

' 由使用者選取範圍，臨時設定列印範圍後匯出，完成後清除設定
Sub SetPrintAreaAndExportToPDF()
    Dim ws          As Worksheet
    Dim rng         As Range
    Dim folderPath  As String
    Dim pdfPath     As String
    Dim origArea    As String

    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要設定為列印範圍的儲存格區域。", vbExclamation, "提示"
        Exit Sub
    End If

    Set rng = Selection

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    pdfPath  = folderPath & "\" & ws.Name & "_CustomPrintArea.pdf"
    origArea = ws.PageSetup.PrintArea

    ' 臨時套用列印範圍
    ws.PageSetup.PrintArea = rng.Address

    On Error GoTo RestoreArea
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "選取範圍已匯出為 PDF：" & vbCrLf & pdfPath, vbInformation, "完成"

RestoreArea:
    If Err.Number <> 0 Then
        MsgBox "匯出失敗：" & Err.Description, vbCritical, "錯誤"
    End If
    ' 還原原始列印範圍設定
    ws.PageSetup.PrintArea = origArea
End Sub
