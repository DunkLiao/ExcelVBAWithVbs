Attribute VB_Name = "SaveSelectedMailAttachments"
Option Explicit
'*************************************************************************************
' 專案名稱: 儲存所選郵件附件範例
' 功能說明: 將 Outlook 目前選取郵件的所有附件存到指定資料夾
'*************************************************************************************

Private Const OL_MAIL_ITEM As Long = 43

Public Sub SaveSelectedMailAttachmentsExample()
    On Error GoTo ErrorHandler

    Dim targetFolder As String
    Dim outlookApp As Object
    Dim selectionItems As Object
    Dim mailItem As Object
    Dim attachmentItem As Object
    Dim savedCount As Long
    Dim filePath As String

    targetFolder = InputBox("請輸入附件儲存資料夾路徑", "儲存附件", "C:\Temp\OutlookAttachments")
    If Len(Trim$(targetFolder)) = 0 Then
        Exit Sub
    End If

    If Right$(targetFolder, 1) <> "\" Then
        targetFolder = targetFolder & "\"
    End If

    EnsureSaveSelectedAttachmentsFolder targetFolder

    Set outlookApp = GetSaveSelectedAttachmentsOutlookApp()
    Set selectionItems = outlookApp.ActiveExplorer.Selection

    For Each mailItem In selectionItems
        If mailItem.Class = OL_MAIL_ITEM Then
            For Each attachmentItem In mailItem.Attachments
                filePath = targetFolder & Format$(Now, "yyyymmdd_hhnnss_") & CleanSaveSelectedAttachmentsFileName(attachmentItem.FileName)
                attachmentItem.SaveAsFile filePath
                savedCount = savedCount + 1
            Next attachmentItem
        End If
    Next mailItem

    MsgBox "附件儲存完成，共儲存 " & CStr(savedCount) & " 個檔案。", vbInformation, "儲存附件"

CleanExit:
    Set attachmentItem = Nothing
    Set mailItem = Nothing
    Set selectionItems = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "儲存附件時發生錯誤：" & Err.Description, vbExclamation, "儲存附件"
    Resume CleanExit
End Sub

Private Function GetSaveSelectedAttachmentsOutlookApp() As Object
    On Error Resume Next

    Set GetSaveSelectedAttachmentsOutlookApp = GetObject(, "Outlook.Application")
    If GetSaveSelectedAttachmentsOutlookApp Is Nothing Then
        Set GetSaveSelectedAttachmentsOutlookApp = CreateObject("Outlook.Application")
    End If
End Function

Private Sub EnsureSaveSelectedAttachmentsFolder(ByVal folderPath As String)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    Set fso = Nothing
End Sub

Private Function CleanSaveSelectedAttachmentsFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    CleanSaveSelectedAttachmentsFileName = fileName

    For Each item In invalidChars
        CleanSaveSelectedAttachmentsFileName = Replace$(CleanSaveSelectedAttachmentsFileName, CStr(item), "_")
    Next item
End Function