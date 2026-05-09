Attribute VB_Name = "BatchProtectWordDocs"
Option Explicit
'*************************************************************************************
'模組名稱: BatchProtectWordDocs
'功能說明: 批次為資料夾內所有 Word 文件設定或解除密碼保護
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 執行 BatchProtectWordDocuments 套用保護
'     或執行 BatchUnprotectWordDocuments 解除保護
'  2. 輸入密碼
'  3. 選擇目標資料夾，程式自動處理所有 .docx
'
'注意事項:
'  - 保護類型為「僅允許填寫表單」（wdAllowOnlyFormFields = 2）
'  - 解除保護時若密碼錯誤，該文件將略過並記錄於結果訊息
'*************************************************************************************

'批次保護 Word 文件（鎖定編輯）
Sub BatchProtectWordDocuments()
    Dim wdApp     As Object
    Dim wdDoc     As Object
    Dim strFolder As String
    Dim strFile   As String
    Dim strPwd    As String
    Dim lngOK     As Long
    Dim lngFail   As Long
    Dim strFailed As String

    On Error GoTo ErrHandler

    strPwd = InputBox("請輸入要設定的保護密碼：", "設定文件保護")
    If strPwd = "" Then
        MsgBox "密碼為空，操作取消。", vbInformation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(4)
        .Title = "請選擇 Word 文件資料夾"
        If .Show <> -1 Then Exit Sub
        strFolder = .SelectedItems(1)
    End With
    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    lngOK = 0
    lngFail = 0

    strFile = Dir(strFolder & "*.docx")
    Do While strFile <> ""
        On Error Resume Next
        Set wdDoc = wdApp.Documents.Open(strFolder & strFile)

        If Err.Number = 0 Then
            '套用保護（wdAllowOnlyFormFields = 2）
            wdDoc.Protect Type:=2, Password:=strPwd, NoReset:=True
            wdDoc.Save
            wdDoc.Close False
            Set wdDoc = Nothing
            lngOK = lngOK + 1
        Else
            lngFail = lngFail + 1
            strFailed = strFailed & strFile & vbCrLf
            Err.Clear
        End If
        On Error GoTo ErrHandler

        strFile = Dir()
    Loop

    wdApp.Quit
    Set wdApp = Nothing

    Dim strMsg As String
    strMsg = "保護設定完成！" & vbCrLf & _
             "成功：" & lngOK & " 個" & vbCrLf & _
             "失敗：" & lngFail & " 個"
    If strFailed <> "" Then strMsg = strMsg & vbCrLf & "失敗清單：" & vbCrLf & strFailed
    MsgBox strMsg, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub

'批次解除 Word 文件保護
Sub BatchUnprotectWordDocuments()
    Dim wdApp     As Object
    Dim wdDoc     As Object
    Dim strFolder As String
    Dim strFile   As String
    Dim strPwd    As String
    Dim lngOK     As Long
    Dim lngFail   As Long
    Dim strFailed As String

    On Error GoTo ErrHandler

    strPwd = InputBox("請輸入現有保護密碼：", "解除文件保護")
    If strPwd = "" Then
        MsgBox "密碼為空，操作取消。", vbInformation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(4)
        .Title = "請選擇 Word 文件資料夾"
        If .Show <> -1 Then Exit Sub
        strFolder = .SelectedItems(1)
    End With
    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    lngOK = 0
    lngFail = 0

    strFile = Dir(strFolder & "*.docx")
    Do While strFile <> ""
        On Error Resume Next
        Set wdDoc = wdApp.Documents.Open(strFolder & strFile)

        If Err.Number = 0 Then
            If wdDoc.ProtectionType <> -1 Then   'wdNoProtection = -1
                wdDoc.Unprotect Password:=strPwd
                If Err.Number = 0 Then
                    wdDoc.Save
                    lngOK = lngOK + 1
                Else
                    lngFail = lngFail + 1
                    strFailed = strFailed & strFile & "（密碼錯誤）" & vbCrLf
                    Err.Clear
                End If
            Else
                lngOK = lngOK + 1  '未保護，直接計數
            End If
            wdDoc.Close False
            Set wdDoc = Nothing
        Else
            lngFail = lngFail + 1
            strFailed = strFailed & strFile & vbCrLf
            Err.Clear
        End If
        On Error GoTo ErrHandler

        strFile = Dir()
    Loop

    wdApp.Quit
    Set wdApp = Nothing

    Dim strMsg As String
    strMsg = "解除保護完成！" & vbCrLf & _
             "成功：" & lngOK & " 個" & vbCrLf & _
             "失敗：" & lngFail & " 個"
    If strFailed <> "" Then strMsg = strMsg & vbCrLf & "失敗清單：" & vbCrLf & strFailed
    MsgBox strMsg, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
