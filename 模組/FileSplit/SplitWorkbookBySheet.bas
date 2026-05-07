Attribute VB_Name = "SplitWorkbookBySheet"
Option Explicit

' ============================================================
' 範例：將活頁簿每個工作表分別存成獨立的 Excel 檔案
' 功能：依工作表名稱逐一另存新檔至指定資料夾
' ============================================================
Sub SplitWorkbookBySheet()
    Dim strFolder   As String
    Dim ws          As Worksheet
    Dim wbNew       As Workbook
    Dim strPath     As String
    Dim intCount    As Integer

    ' --- 選擇儲存資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    intCount = 0

    ' --- 逐一複製工作表並另存 ---
    For Each ws In ThisWorkbook.Worksheets
        ws.Copy
        Set wbNew = ActiveWorkbook
        strPath = strFolder & ws.Name & ".xlsx"
        wbNew.SaveAs Filename:=strPath, FileFormat:=xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
        intCount = intCount + 1
    Next ws

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "已成功切割 " & intCount & " 個工作表，輸出至：" & strFolder, vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
