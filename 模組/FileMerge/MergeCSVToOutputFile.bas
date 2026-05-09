Attribute VB_Name = "MergeCSVToOutputFile"
Option Explicit

' ============================================================
' 範例：將資料夾內所有 CSV 檔案合併輸出為單一 CSV 檔案
' 功能：不匯入 Excel，直接以文字串流方式讀取並輸出
'       第一個檔案保留標題列，後續檔案跳過標題
' ============================================================

Sub MergeCSVFilesToSingleFile()
    On Error GoTo ErrHandler

    Dim strFolder   As String
    Dim strOutFile  As String
    Dim strFile     As String
    Dim intIn       As Integer
    Dim intOut      As Integer
    Dim strLine     As String
    Dim blnFirst    As Boolean
    Dim lngFiles    As Long
    Dim lngLines    As Long

    ' --- 選擇資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 CSV 檔案的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
    strOutFile = strFolder & "MergedOutput.csv"

    ' 防止輸出檔已存在時重複處理
    If Dir(strOutFile) <> "" Then
        Dim intResp As Integer
        intResp = MsgBox("輸出檔案已存在：" & strOutFile & vbCrLf & "是否覆寫？", _
                         vbYesNo + vbQuestion, "確認覆寫")
        If intResp = vbNo Then Exit Sub
    End If

    intOut = FreeFile
    Open strOutFile For Output As #intOut

    blnFirst = True
    lngFiles = 0
    lngLines = 0

    strFile = Dir(strFolder & "*.csv")
    If strFile = "" Then
        MsgBox "找不到任何 CSV 檔案。", vbExclamation, "警告"
        Close #intOut
        Kill strOutFile
        Exit Sub
    End If

    Do While strFile <> ""
        ' 略過輸出檔本身
        If StrComp(strFile, "MergedOutput.csv", vbTextCompare) <> 0 Then
            intIn = FreeFile
            Open strFolder & strFile For Input As #intIn

            Dim blnIsFirstLine As Boolean
            blnIsFirstLine = True

            Do While Not EOF(intIn)
                Line Input #intIn, strLine
                If blnIsFirstLine Then
                    If blnFirst Then
                        ' 第一個檔案的標題列直接寫入
                        Print #intOut, strLine
                        lngLines = lngLines + 1
                        blnFirst = False
                    End If
                    ' 後續檔案跳過標題列
                    blnIsFirstLine = False
                Else
                    Print #intOut, strLine
                    lngLines = lngLines + 1
                End If
            Loop

            Close #intIn
            lngFiles = lngFiles + 1
        End If
        strFile = Dir()
    Loop

    Close #intOut

    MsgBox "CSV 合併完成！" & vbCrLf & _
           "共合併 " & lngFiles & " 個檔案，輸出 " & lngLines & " 列。" & vbCrLf & _
           "輸出路徑：" & strOutFile, vbInformation, "完成"
    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #intIn
    Close #intOut
    On Error GoTo 0
    MsgBox "合併 CSV 時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub