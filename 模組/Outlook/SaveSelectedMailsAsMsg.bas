Attribute VB_Name = "SaveSelectedMailsAsMsg"
Option Explicit
'*************************************************************************************
' 專案名稱: 將所選郵件另存為 MSG 範例
' 功能說明: 將 Outlook 目前選取郵件逐一另存成 .msg 檔
'*************************************************************************************

Private Const OL_MAIL_ITEM As Long = 43
Private Const OL_MSG As Long = 3

Public Sub SaveSelectedMailsAsMsgExample()
    On Error GoTo ErrorHandler

    Dim targetFolder As String
    Dim outlookApp As Object
    Dim selectionItems As Object
    Dim mailItem As Object
    Dim savedCount As Long
    Dim fileName As String

    targetFolder = InputBox("請輸入 MSG 儲存資料夾路徑", "另存郵件", "C:\Temp\OutlookMsg")
    If Len(Trim$(targetFolder)) = 0 Then
        Exit Sub
    End If

    If Right$(targetFolder, 1) <> "\" Then
        targetFolder = targetFolder & "\"
    End If

    EnsureSaveSelectedMsgFolder targetFolder

    Set outlookApp = GetSaveSelectedMsgOutlookApp()
    Set selectionItems = outlookApp.ActiveExplorer.Selection

    For Each mailItem In selectionItems
        If mailItem.Class = OL_MAIL_ITEM Then
            fileName = targetFolder & Format$(mailItem.ReceivedTime, "yyyymmdd_hhnnss_") & CleanSaveSelectedMsgFileName(mailItem.Subject) & ".msg"
            mailItem.SaveAs fileName, OL_MSG
            savedCount = savedCount + 1
        End If
    Next mailItem

    MsgBox "郵件另存完成，共儲存 " & CStr(savedCount) & " 封。", vbInformation, "另存郵件"

CleanExit:
    Set mailItem = Nothing
    Set selectionItems = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "另存郵件時發生錯誤：" & Err.Description, vbExclamation, "另存郵件"
    Resume CleanExit
End Sub

Private Function GetSaveSelectedMsgOutlookApp() As Object
    On Error Resume Next

    Set GetSaveSelectedMsgOutlookApp = GetObject(, "Outlook.Application")
    If GetSaveSelectedMsgOutlookApp Is Nothing Then
        Set GetSaveSelectedMsgOutlookApp = CreateObject("Outlook.Application")
    End If
End Function

Private Sub EnsureSaveSelectedMsgFolder(ByVal folderPath As String)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    Set fso = Nothing
End Sub

Private Function CleanSaveSelectedMsgFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    CleanSaveSelectedMsgFileName = Left$(fileName, 80)

    For Each item In invalidChars
        CleanSaveSelectedMsgFileName = Replace$(CleanSaveSelectedMsgFileName, CStr(item), "_")
    Next item

    If Len(Trim$(CleanSaveSelectedMsgFileName)) = 0 Then
        CleanSaveSelectedMsgFileName = "NoSubject"
    End If
End Function