Attribute VB_Name = "ListUnreadMailsToImmediate"
Option Explicit
'*************************************************************************************
' 專案名稱: 列出未讀郵件範例
' 功能說明: 將目前 Outlook 資料夾中的未讀郵件資訊輸出到立即視窗
'*************************************************************************************

Private Const OL_MAIL_ITEM As Long = 43

Public Sub ListUnreadMailsToImmediateExample()
    On Error GoTo ErrorHandler

    Dim outlookApp As Object
    Dim currentFolder As Object
    Dim mailItem As Object
    Dim index As Long
    Dim unreadCount As Long

    Set outlookApp = GetListUnreadMailsOutlookApp()
    Set currentFolder = outlookApp.ActiveExplorer.CurrentFolder

    Debug.Print String$(60, "-")
    Debug.Print "未讀郵件清單：" & currentFolder.Name
    Debug.Print String$(60, "-")

    For index = 1 To currentFolder.Items.Count
        Set mailItem = currentFolder.Items.Item(index)
        If mailItem.Class = OL_MAIL_ITEM Then
            If mailItem.UnRead = True Then
                Debug.Print Format$(mailItem.ReceivedTime, "yyyy/mm/dd hh:nn") & " | " & mailItem.SenderName & " | " & mailItem.Subject
                unreadCount = unreadCount + 1
            End If
        End If
    Next index

    MsgBox "已列出 " & CStr(unreadCount) & " 封未讀郵件，請查看 VBA 立即視窗。", vbInformation, "未讀郵件"

CleanExit:
    Set mailItem = Nothing
    Set currentFolder = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "列出未讀郵件時發生錯誤：" & Err.Description, vbExclamation, "未讀郵件"
    Resume CleanExit
End Sub

Private Function GetListUnreadMailsOutlookApp() As Object
    On Error Resume Next

    Set GetListUnreadMailsOutlookApp = GetObject(, "Outlook.Application")
    If GetListUnreadMailsOutlookApp Is Nothing Then
        Set GetListUnreadMailsOutlookApp = CreateObject("Outlook.Application")
    End If
End Function