Attribute VB_Name = "MergeExcelWithSheetProtection"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithSheetProtection
'功能說明: 合併多個Excel檔案時，自動解除來源工作表保護，合併後重新保護目的表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestMergeWithProtection()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇要合併的Excel檔案資料夾"
        If .Show = True Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
    End With
    Call MergeExcelWithSheetProtection(folderPath, "", "合併結果")
End Sub

Sub MergeExcelWithSheetProtection(ByVal folderPath As String, _
                                   ByVal password As String, _
                                   ByVal targetSheetName As String)
    Dim fso        As Object
    Dim folder     As Object
    Dim fileItem   As Object
    Dim srcWb      As Workbook
    Dim ws         As Worksheet
    Dim destWs     As Worksheet
    Dim destRow    As Long
    Dim hasHeader  As Boolean
    Dim lastRow    As Long
    Dim lastCol    As Long
    Dim ext        As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        MsgBox "資料夾不存在：" & folderPath, vbExclamation, "錯誤"
        Exit Sub
    End If

    On Error Resume Next
    Set destWs = ThisWorkbook.Worksheets(targetSheetName)
    On Error GoTo 0
    If destWs Is Nothing Then
        Set destWs = ThisWorkbook.Worksheets.Add
        destWs.Name = targetSheetName
    End If
    destWs.Cells.Clear
    destRow = 1
    hasHeader = False

    Set folder = fso.GetFolder(folderPath)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each fileItem In folder.Files
        ext = LCase(fso.GetExtensionName(fileItem.Name))
        If ext = "xlsx" Or ext = "xls" Then
            Set srcWb = Workbooks.Open(fileItem.Path, ReadOnly:=True)
            For Each ws In srcWb.Worksheets
                On Error Resume Next
                ws.Unprotect password
                On Error GoTo 0
                lastRow = ws.UsedRange.Rows.Count
                lastCol = ws.UsedRange.Columns.Count
                If lastRow > 0 And lastCol > 0 Then
                    If Not hasHeader Then
                        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy _
                            Destination:=destWs.Cells(destRow, 1)
                        destRow = destRow + lastRow
                        hasHeader = True
                    Else
                        If lastRow > 1 Then
                            ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy _
                                Destination:=destWs.Cells(destRow, 1)
                            destRow = destRow + lastRow - 1
                        End If
                    End If
                End If
                If password <> "" Then ws.Protect password
            Next ws
            srcWb.Close SaveChanges:=False
        End If
    Next fileItem

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    destWs.Columns.AutoFit
    MsgBox "合併完成！共 " & destRow - 1 & " 列資料。", vbInformation, "完成"
End Sub
