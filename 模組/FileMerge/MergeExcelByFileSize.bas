Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelByFileSize
'功能說明: 依照檔案大小由小到大排序後，依序合併指定資料夾內的所有 Excel 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub MergeExcelByFileSize()
    Dim folderPath As String
    Dim outWb As Workbook
    Dim outWs As Worksheet
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim fso As Object
    Dim folder As Object
    Dim fileObj As Object
    Dim destRow As Long
    Dim hasHeader As Boolean
    Dim fileCount As Long
    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim maxLastCol As Long
    Dim tempName As String
    Dim tempSize As Long
    Dim fileList() As String
    Dim fileSize() As Long

    On Error GoTo ErrHandler

    ' 選擇資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇要合併的 Excel 檔案資料夾"
        If .Show = False Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' 列舉 Excel 檔案
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    fileCount = 0
    For Each fileObj In folder.Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(fileObj.Name))
        If ext = "xlsx" Or ext = "xls" Or ext = "xlsm" Then
            fileCount = fileCount + 1
            ReDim Preserve fileList(1 To fileCount)
            ReDim Preserve fileSize(1 To fileCount)
            fileList(fileCount) = fileObj.Path
            fileSize(fileCount) = fileObj.Size
        End If
    Next fileObj

    If fileCount = 0 Then
        MsgBox "資料夾中找不到 Excel 檔案！", vbExclamation, "提示"
        Exit Sub
    End If

    ' 依檔案大小排序（由小到大，氣泡排序）
    For i = 1 To fileCount - 1
        For j = i + 1 To fileCount
            If fileSize(i) > fileSize(j) Then
                tempSize = fileSize(i): fileSize(i) = fileSize(j): fileSize(j) = tempSize
                tempName = fileList(i): fileList(i) = fileList(j): fileList(j) = tempName
            End If
        Next j
    Next i

    ' 建立輸出活頁簿
    Set outWb = Workbooks.Add
    Set outWs = outWb.Sheets(1)
    outWs.Name = "合併結果"
    destRow = 1
    hasHeader = False
    maxLastCol = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 依序合併每個檔案
    For i = 1 To fileCount
        On Error Resume Next
        Set srcWb = Workbooks.Open(fileList(i), ReadOnly:=True)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextFile
        End If
        On Error GoTo ErrHandler

        Set srcWs = srcWb.Sheets(1)
        lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
        lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

        If lastRow < 1 Or lastCol < 1 Then
            srcWb.Close SaveChanges:=False
            GoTo NextFile
        End If

        If lastCol > maxLastCol Then maxLastCol = lastCol

        If Not hasHeader Then
            ' 第一個檔案：含標題整行複製
            srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol)).Copy _
                Destination:=outWs.Cells(destRow, 1)
            ' 填入來源欄資訊
            outWs.Cells(1, lastCol + 1).Value = "來源檔案（大小）"
            Dim r As Long
            For r = 2 To lastRow
                outWs.Cells(r, lastCol + 1).Value = _
                    fso.GetFileName(fileList(i)) & " (" & Format(fileSize(i), "#,##0") & " bytes)"
            Next r
            destRow = destRow + lastRow
            hasHeader = True
        Else
            ' 後續檔案：跳過第一列標題
            If lastRow > 1 Then
                srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(lastRow, lastCol)).Copy _
                    Destination:=outWs.Cells(destRow, 1)
                Dim r2 As Long
                For r2 = destRow To destRow + lastRow - 2
                    outWs.Cells(r2, maxLastCol + 1).Value = _
                        fso.GetFileName(fileList(i)) & " (" & Format(fileSize(i), "#,##0") & " bytes)"
                Next r2
                destRow = destRow + lastRow - 1
            End If
        End If

        srcWb.Close SaveChanges:=False
        GoTo NextFile

NextFile:
        Set srcWb = Nothing
    Next i

    outWs.Columns.AutoFit

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "依檔案大小合併完成！共合併 " & fileCount & " 個檔案。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
