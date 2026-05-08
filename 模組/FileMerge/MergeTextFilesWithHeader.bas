Option Explicit

Private Const MsoFileDialogFolderPicker As Long = 4

' 合併資料夾內所有 TXT 檔，並只保留第一個檔案的標題列。
Public Sub MergeTextFilesWithHeaderExample()
    On Error GoTo ErrHandler

    Dim folderPath As String
    Dim outputPath As String

    folderPath = PickMergeFolder()
    If Len(folderPath) = 0 Then Exit Sub

    outputPath = folderPath & "\MergedTextWithHeader.txt"
    Call MergeTextFilesWithHeader(folderPath, outputPath)

    MsgBox "文字檔合併完成：" & outputPath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "合併文字檔失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub MergeTextFilesWithHeader(ByVal folderPath As String, ByVal outputPath As String)
    Dim fileName As String
    Dim inputFile As Integer
    Dim outputFile As Integer
    Dim lineText As String
    Dim isFirstFile As Boolean
    Dim isFirstLine As Boolean

    outputFile = FreeFile
    Open outputPath For Output As #outputFile

    isFirstFile = True
    fileName = Dir(folderPath & "\*.txt")
    Do While Len(fileName) > 0
        If StrComp(fileName, "MergedTextWithHeader.txt", vbTextCompare) <> 0 Then
            inputFile = FreeFile
            Open folderPath & "\" & fileName For Input As #inputFile
            isFirstLine = True

            Do While Not EOF(inputFile)
                Line Input #inputFile, lineText
                If isFirstFile Or Not isFirstLine Then
                    Print #outputFile, lineText
                End If
                isFirstLine = False
            Loop

            Close #inputFile
            isFirstFile = False
        End If
        fileName = Dir()
    Loop

    Close #outputFile
End Sub

Private Function PickMergeFolder() As String
    With Application.FileDialog(MsoFileDialogFolderPicker)
        .Title = "請選擇要合併 TXT 檔的資料夾"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickMergeFolder = .SelectedItems(1)
        End If
    End With
End Function