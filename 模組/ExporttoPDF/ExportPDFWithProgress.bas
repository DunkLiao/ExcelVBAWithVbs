Attribute VB_Name = "ExportPDFWithProgress"
Option Explicit

' ============================================================
' 範例：逐一匯出工作表為 PDF，並在狀態列即時顯示進度文字
' 功能：每匯出一張工作表即更新 Excel 狀態列，完成後自動還原
' ============================================================

Sub ExportPDFWithProgress()
    Dim folderPath  As String
    Dim ws          As Worksheet
    Dim pdfPath     As String
    Dim total       As Integer
    Dim current     As Integer
    Dim errCount    As Integer

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    total    = ThisWorkbook.Worksheets.Count
    current  = 0
    errCount = 0

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        current = current + 1
        ' 更新狀態列顯示目前進度
        Application.StatusBar = "正在匯出第 " & current & " / " & total & _
                                " 張：" & ws.Name & "　請稍候..."

        pdfPath = folderPath & "\" & ws.Name & ".pdf"

        On Error Resume Next
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=pdfPath, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        If Err.Number <> 0 Then errCount = errCount + 1
        Err.Clear
        On Error GoTo 0
    Next ws

    Application.ScreenUpdating = True
    Application.StatusBar = False  ' 還原狀態列為預設文字

    MsgBox "匯出完成！" & vbCrLf & _
           "成功：" & (total - errCount) & " 張" & vbCrLf & _
           "失敗：" & errCount & " 張", vbInformation, "完成"
End Sub
