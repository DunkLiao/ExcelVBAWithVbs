Attribute VB_Name = "MergeWordDocuments"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWordDocuments
'功能說明: 將指定資料夾內的所有 Word 文件合併為單一文件
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 執行 MergeAllWordDocumentsInFolder
'  2. 選擇包含多個 .docx 文件的資料夾
'  3. 程式依檔名排序後逐一合併，每份文件後自動插入分頁符號
'  4. 完成後提示儲存合併後文件
'
'注意事項:
'  - 合併後的文件為全新文件，原始文件不受影響
'  - 需啟用 Microsoft Word Object Library 參考
'*************************************************************************************

'合併資料夾內所有 Word 文件
Sub MergeAllWordDocumentsInFolder()
    Dim wdApp       As Object
    Dim wdMerged    As Object
    Dim wdSrc       As Object
    Dim strFolder   As String
    Dim strFile     As String
    Dim strSavePath As String
    Dim strTemp     As String
    Dim colFiles    As Collection
    Dim varFile     As Variant
    Dim lngCount    As Long
    Dim arrFiles()  As String
    Dim k           As Long
    Dim m           As Long

    On Error GoTo ErrHandler

    '選擇目標資料夾（msoFileDialogFolderPicker = 4）
    With Application.FileDialog(4)
        .Title = "請選擇要合併的 Word 文件資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    '收集所有 .docx 檔案
    Set colFiles = New Collection
    strFile = Dir(strFolder & "*.docx")
    Do While strFile <> ""
        colFiles.Add strFolder & strFile
        strFile = Dir()
    Loop

    If colFiles.Count = 0 Then
        MsgBox "所選資料夾中找不到 .docx 文件！", vbExclamation, "提示"
        Exit Sub
    End If

    '將 Collection 轉為陣列以便排序
    ReDim arrFiles(1 To colFiles.Count)
    k = 1
    For Each varFile In colFiles
        arrFiles(k) = CStr(varFile)
        k = k + 1
    Next varFile

    '氣泡排序（依檔名升冪）
    For k = 1 To UBound(arrFiles) - 1
        For m = 1 To UBound(arrFiles) - k
            If arrFiles(m) > arrFiles(m + 1) Then
                strTemp = arrFiles(m)
                arrFiles(m) = arrFiles(m + 1)
                arrFiles(m + 1) = strTemp
            End If
        Next m
    Next k

    '建立 Word 物件
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    '建立合併目標文件
    Set wdMerged = wdApp.Documents.Add
    lngCount = 0

    '逐一插入各文件內容
    For k = 1 To UBound(arrFiles)
        Set wdSrc = wdApp.Documents.Open(arrFiles(k))

        '複製全文內容至合併文件末端
        wdSrc.Content.Copy
        wdMerged.Bookmarks("\EndOfDoc").Range.Paste

        '插入分頁符號（最後一份除外，wdPageBreak = 7）
        If k < UBound(arrFiles) Then
            wdMerged.Bookmarks("\EndOfDoc").Range.InsertBreak 7
        End If

        wdSrc.Close False
        Set wdSrc = Nothing
        lngCount = lngCount + 1
    Next k

    '顯示合併文件
    wdApp.Visible = True

    '提示儲存路徑
    strSavePath = Application.GetSaveAsFilename( _
        InitialFileName:="合併文件.docx", _
        FileFilter:="Word 文件 (*.docx), *.docx")

    If strSavePath <> "False" Then
        wdMerged.SaveAs2 strSavePath, 16
        MsgBox "合併完成！" & vbCrLf & _
               "共合併 " & lngCount & " 個文件。" & vbCrLf & _
               "已儲存至：" & strSavePath, vbInformation, "完成"
    Else
        MsgBox "已取消儲存，合併文件仍開啟中。", vbInformation, "提示"
    End If

    Set wdMerged = Nothing
    Set wdApp = Nothing
    Set colFiles = Nothing
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdSrc Is Nothing Then wdSrc.Close False
    If Not wdMerged Is Nothing Then wdMerged.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdSrc = Nothing
    Set wdMerged = Nothing
    Set wdApp = Nothing
End Sub
