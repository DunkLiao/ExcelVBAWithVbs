Attribute VB_Name = "ExportPDFWithAutoFileName"
Option Explicit

' ============================================================
' 範例：自動以儲存格值與時間戳記產生 PDF 檔名後匯出
' 功能：讀取 A1 儲存格的文字作為檔名前綴，自動加上日期時間後存檔
' ============================================================

Sub ExportPDFWithAutoFileName()
    Dim ws          As Worksheet
    Dim folderPath  As String
    Dim cellValue   As String
    Dim pdfPath     As String

    Set ws = ActiveSheet

    ' 取得 A1 儲存格內容作為檔名前綴
    cellValue = Trim(ws.Range("A1").Value)
    If Len(cellValue) = 0 Then cellValue = ws.Name

    ' 選擇輸出資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    ' 組合完整路徑：資料夾 + A1內容 + 時間戳記
    pdfPath = folderPath & "\" & cellValue & "_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"

    On Error GoTo ErrHandler
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "PDF 已成功匯出！" & vbCrLf & "儲存路徑：" & pdfPath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "匯出失敗！" & vbCrLf & "錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
