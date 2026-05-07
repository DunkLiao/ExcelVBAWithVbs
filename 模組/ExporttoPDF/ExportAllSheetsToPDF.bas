Attribute VB_Name = "ExportAllSheetsToPDF"
Option Explicit

' ============================================================
' 範例：將活頁簿所有工作表合併匯出為單一 PDF 檔案
' 功能：選擇輸出路徑後，將所有工作表一次列印為 PDF
' ============================================================
Sub ExportAllSheetsToPDF()
    Dim strPath     As String
    Dim wsArr()     As String
    Dim i           As Integer

    ' --- 選擇儲存路徑 ---
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請選擇 PDF 儲存路徑"
        .FilterIndex = 2
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation
            Exit Sub
        End If
        strPath = .SelectedItems(1)
    End With

    If Right(LCase(strPath), 4) <> ".pdf" Then strPath = strPath & ".pdf"

    On Error GoTo ErrHandler

    ' --- 收集所有工作表名稱 ---
    ReDim wsArr(1 To ThisWorkbook.Sheets.Count)
    For i = 1 To ThisWorkbook.Sheets.Count
        wsArr(i) = ThisWorkbook.Sheets(i).Name
    Next i

    ' --- 選取所有工作表並匯出 ---
    ThisWorkbook.Sheets(wsArr).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=strPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False

    ' --- 取消多工作表選取 ---
    ThisWorkbook.Sheets(1).Select

    MsgBox "所有工作表已匯出為 PDF：" & Chr(10) & strPath, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
