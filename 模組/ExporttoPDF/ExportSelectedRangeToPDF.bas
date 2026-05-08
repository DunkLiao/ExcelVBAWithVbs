Option Explicit

Private Const MsoFileDialogFolderPicker As Long = 4

' 將目前選取範圍轉存成 PDF。
Public Sub ExportSelectedRangeToPDFExample()
    On Error GoTo ErrHandler

    Dim selectedRange As Range
    Dim folderPath As String
    Dim pdfPath As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要轉存的儲存格範圍。", vbExclamation, "提醒"
        Exit Sub
    End If

    Set selectedRange = Selection
    folderPath = PickPdfFolder()
    If Len(folderPath) = 0 Then Exit Sub

    pdfPath = folderPath & "\SelectedRange_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    selectedRange.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    MsgBox "PDF 已轉存完成：" & pdfPath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "轉存 PDF 失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function PickPdfFolder() As String
    With Application.FileDialog(MsoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickPdfFolder = .SelectedItems(1)
        End If
    End With
End Function