Attribute VB_Name = "ExportPDFBySheetCondition"
Option Explicit

' ============================================================
' 範例：依條件篩選工作表後批次匯出 PDF
' 功能：支援只匯出可見工作表，或依名稱關鍵字篩選後匯出
' ============================================================

' 只匯出所有可見（非隱藏）的工作表為個別 PDF
Sub ExportVisibleSheetsToPDF()
    Dim folderPath  As String
    Dim ws          As Worksheet
    Dim pdfPath     As String
    Dim totalCount  As Integer

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    totalCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            pdfPath = folderPath & "\" & ws.Name & ".pdf"
            On Error Resume Next
            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=pdfPath, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
            If Err.Number = 0 Then totalCount = totalCount + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next ws

    Application.ScreenUpdating = True
    MsgBox "已匯出所有可見工作表為 PDF，共 " & totalCount & " 個。", vbInformation, "完成"
End Sub

' 依工作表名稱關鍵字篩選，符合者匯出為 PDF
Sub ExportSheetsByKeywordToPDF()
    Dim folderPath  As String
    Dim ws          As Worksheet
    Dim pdfPath     As String
    Dim keyword     As String
    Dim totalCount  As Integer

    keyword = InputBox("請輸入工作表名稱關鍵字（模糊比對）：", "關鍵字篩選匯出")
    If Trim(keyword) = "" Then
        MsgBox "未輸入關鍵字，操作取消。", vbInformation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    totalCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, keyword, vbTextCompare) > 0 Then
            pdfPath = folderPath & "\" & ws.Name & ".pdf"
            On Error Resume Next
            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=pdfPath, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
            If Err.Number = 0 Then totalCount = totalCount + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next ws

    Application.ScreenUpdating = True

    If totalCount = 0 Then
        MsgBox "沒有符合關鍵字「" & keyword & "」的工作表。", vbExclamation, "提示"
    Else
        MsgBox "已匯出名稱含「" & keyword & "」的工作表，共 " & totalCount & " 個 PDF。", vbInformation, "完成"
    End If
End Sub
