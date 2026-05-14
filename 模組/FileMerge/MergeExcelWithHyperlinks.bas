Attribute VB_Name = "MergeExcelWithHyperlinks"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithHyperlinks
'功能說明: 合併多個 Excel 檔案並保留儲存格超連結的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點：選擇資料夾後合併所有 Excel 並保留超連結
Sub TestMergeExcelWithHyperlinks()
    Dim folderPath As String
    folderPath = GetHyperlinkFolderPath()
    If folderPath = "" Then
        MsgBox "未選擇資料夾，作業取消。", vbInformation, "取消"
        Exit Sub
    End If
    Call MergeExcelKeepHyperlinks(folderPath)
End Sub

' 合併指定資料夾下所有 .xlsx 並保留超連結
' folderPath: 來源資料夾路徑
Sub MergeExcelKeepHyperlinks(ByVal folderPath As String)
    On Error GoTo ErrorHandler

    Dim fso       As Object
    Dim srcFolder As Object
    Dim f         As Object
    Dim srcWb     As Workbook
    Dim srcWs     As Worksheet
    Dim destWs    As Worksheet
    Dim destRow   As Long
    Dim srcLast   As Long
    Dim srcLastCol As Long
    Dim hLink     As Hyperlink
    Dim relRow    As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        MsgBox "資料夾不存在：" & folderPath, vbExclamation, "錯誤"
        Exit Sub
    End If

    Set destWs = GetOrCreateHyperlinkSheet("合併結果")
    destWs.Cells.Clear
    destRow = 1

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set srcFolder = fso.GetFolder(folderPath)
    For Each f In srcFolder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "xlsx" Then
            Set srcWb = Workbooks.Open(f.Path, ReadOnly:=True)
            Set srcWs = srcWb.Worksheets(1)

            srcLast = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
            srcLastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

            If srcLast >= 1 And srcLastCol >= 1 Then
                srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(srcLast, srcLastCol)).Copy
                destWs.Cells(destRow, 1).PasteSpecial Paste:=xlPasteValues
                destWs.Cells(destRow, 1).PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False

                For Each hLink In srcWs.Hyperlinks
                    relRow = hLink.Range.Row - 1 + destRow
                    destWs.Hyperlinks.Add _
                        Anchor:=destWs.Cells(relRow, hLink.Range.Column), _
                        Address:=hLink.Address, _
                        SubAddress:=hLink.SubAddress, _
                        TextToDisplay:=hLink.TextToDisplay, _
                        ScreenTip:=hLink.ScreenTip
                Next hLink

                destRow = destRow + srcLast
            End If

            srcWb.Close SaveChanges:=False
        End If
    Next f

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "合併完成（含超連結）！共 " & destRow - 1 & " 列資料。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得資料夾路徑
Private Function GetHyperlinkFolderPath() As String
    Dim dialog As Object
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.Title = "請選擇要合併的 Excel 檔案資料夾"
    If dialog.Show = -1 Then
        GetHyperlinkFolderPath = dialog.SelectedItems(1)
    Else
        GetHyperlinkFolderPath = ""
    End If
End Function

' 取得或建立工作表
Private Function GetOrCreateHyperlinkSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateHyperlinkSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateHyperlinkSheet Is Nothing Then
        Set GetOrCreateHyperlinkSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateHyperlinkSheet.Name = sheetName
    End If
End Function
