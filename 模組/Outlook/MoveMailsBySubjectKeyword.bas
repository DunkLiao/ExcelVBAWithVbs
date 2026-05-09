Attribute VB_Name = "MoveMailsBySubjectKeyword"
Option Explicit
'*************************************************************************************
' 專案名稱: 依主旨關鍵字移動郵件範例
' 功能說明: 將目前資料夾中主旨含指定關鍵字的郵件移到子資料夾
'*************************************************************************************

Private Const OL_MAIL_ITEM As Long = 43

Public Sub MoveMailsBySubjectKeywordExample()
    On Error GoTo ErrorHandler

    Dim keyword As String
    Dim targetFolderName As String
    Dim outlookApp As Object
    Dim sourceFolder As Object
    Dim targetFolder As Object
    Dim mailItem As Object
    Dim index As Long
    Dim movedCount As Long

    keyword = InputBox("請輸入主旨關鍵字", "移動郵件")
    If Len(Trim$(keyword)) = 0 Then
        Exit Sub
    End If

    targetFolderName = InputBox("請輸入目的子資料夾名稱", "移動郵件", "已分類")
    If Len(Trim$(targetFolderName)) = 0 Then
        Exit Sub
    End If

    Set outlookApp = GetMoveMailsBySubjectOutlookApp()
    Set sourceFolder = outlookApp.ActiveExplorer.CurrentFolder
    Set targetFolder = GetMoveMailsBySubjectSubFolder(sourceFolder, targetFolderName)

    For index = sourceFolder.Items.Count To 1 Step -1
        Set mailItem = sourceFolder.Items.Item(index)
        If mailItem.Class = OL_MAIL_ITEM Then
            If InStr(1, mailItem.Subject, keyword, vbTextCompare) > 0 Then
                mailItem.Move targetFolder
                movedCount = movedCount + 1
            End If
        End If
    Next index

    MsgBox "移動完成，共移動 " & CStr(movedCount) & " 封郵件。", vbInformation, "移動郵件"

CleanExit:
    Set mailItem = Nothing
    Set targetFolder = Nothing
    Set sourceFolder = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "移動郵件時發生錯誤：" & Err.Description, vbExclamation, "移動郵件"
    Resume CleanExit
End Sub

Private Function GetMoveMailsBySubjectOutlookApp() As Object
    On Error Resume Next

    Set GetMoveMailsBySubjectOutlookApp = GetObject(, "Outlook.Application")
    If GetMoveMailsBySubjectOutlookApp Is Nothing Then
        Set GetMoveMailsBySubjectOutlookApp = CreateObject("Outlook.Application")
    End If
End Function

Private Function GetMoveMailsBySubjectSubFolder(ByVal parentFolder As Object, ByVal folderName As String) As Object
    On Error Resume Next

    Set GetMoveMailsBySubjectSubFolder = parentFolder.Folders.Item(folderName)
    On Error GoTo 0

    If GetMoveMailsBySubjectSubFolder Is Nothing Then
        Set GetMoveMailsBySubjectSubFolder = parentFolder.Folders.Add(folderName)
    End If
End Function