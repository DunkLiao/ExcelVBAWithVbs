Attribute VB_Name = "MergeJSONFiles"
Option Explicit

' ============================================================
' 範例：合併同一資料夾下所有 .json 文字檔
' 功能：逐一讀取每個 JSON 檔案（假設為 JSON array），
'       去除首尾的 [ ] 後，合併成一個大型 JSON array 輸出
' ============================================================

Sub MergeJSONFilesInFolder()
    On Error GoTo ErrHandler

    Dim strFolder   As String
    Dim strFile     As String
    Dim strOutFile  As String
    Dim intIn       As Integer
    Dim intOut      As Integer
    Dim strLine     As String
    Dim strContent  As String
    Dim blnFirst    As Boolean
    Dim lngCount    As Long

    ' --- 選擇資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 JSON 檔案的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    strOutFile = strFolder & "MergedOutput.json"
    intOut = FreeFile
    Open strOutFile For Output As #intOut
    Print #intOut, "["

    blnFirst = True
    lngCount = 0
    strFile = Dir(strFolder & "*.json")

    If strFile = "" Then
        MsgBox "找不到任何 JSON 檔案。", vbExclamation, "警告"
        Close #intOut
        Kill strOutFile
        Exit Sub
    End If

    Do While strFile <> ""
        ' 略過輸出檔本身
        If StrComp(strFile, "MergedOutput.json", vbTextCompare) <> 0 Then
            strContent = ""
            intIn = FreeFile
            Open strFolder & strFile For Input As #intIn
            Do While Not EOF(intIn)
                Line Input #intIn, strLine
                strContent = strContent & strLine & vbCrLf
            Loop
            Close #intIn

            ' 去除最外層 [ ] (假設為合法 JSON array)
            strContent = Trim(strContent)
            If Left(strContent, 1) = "[" Then
                strContent = Mid(strContent, 2)
            End If
            If Right(strContent, 1) = "]" Then
                strContent = Left(strContent, Len(strContent) - 1)
            End If
            strContent = Trim(strContent)

            If Len(strContent) > 0 Then
                If Not blnFirst Then
                    Print #intOut, ","
                End If
                Print #intOut, strContent
                blnFirst = False
                lngCount = lngCount + 1
            End If
        End If
        strFile = Dir()
    Loop

    Print #intOut, "]"
    Close #intOut

    MsgBox "JSON 合併完成！共合併 " & lngCount & " 個檔案。" & vbCrLf & _
           "輸出路徑：" & strOutFile, vbInformation, "完成"
    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #intOut
    On Error GoTo 0
    MsgBox "合併 JSON 時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub