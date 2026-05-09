Attribute VB_Name = "SendWorkbookAsAttachment"
Option Explicit
'*************************************************************************************
' 專案名稱: 寄送目前活頁簿範例
' 功能說明: 從 Excel 將目前活頁簿存檔後附加到 Outlook 郵件
'*************************************************************************************

Private Const OL_MAIL_ITEM As Long = 0

Public Sub SendWorkbookAsAttachmentExample()
    On Error GoTo ErrorHandler

    Dim outlookApp As Object
    Dim mailItem As Object
    Dim workbookPath As String
    Dim recipientText As String

    recipientText = InputBox("請輸入收件者信箱，多人請用分號分隔", "寄送活頁簿")
    If Len(Trim$(recipientText)) = 0 Then
        Exit Sub
    End If

    workbookPath = ActiveWorkbook.FullName
    If Len(workbookPath) = 0 Then
        MsgBox "請先儲存目前活頁簿，再執行此範例。", vbExclamation, "寄送活頁簿"
        Exit Sub
    End If

    ActiveWorkbook.Save

    Set outlookApp = GetSendWorkbookAttachmentOutlookApp()
    Set mailItem = outlookApp.CreateItem(OL_MAIL_ITEM)

    With mailItem
        .To = recipientText
        .Subject = "活頁簿附件 - " & ActiveWorkbook.Name
        .Body = "您好，請參考附件活頁簿。" & vbCrLf & vbCrLf & "此郵件由 VBA 範例建立，請確認後再寄出。"
        .Attachments.Add workbookPath
        .Display
    End With

CleanExit:
    Set mailItem = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "建立郵件時發生錯誤：" & Err.Description, vbExclamation, "寄送活頁簿"
    Resume CleanExit
End Sub

Private Function GetSendWorkbookAttachmentOutlookApp() As Object
    On Error Resume Next

    Set GetSendWorkbookAttachmentOutlookApp = GetObject(, "Outlook.Application")
    If GetSendWorkbookAttachmentOutlookApp Is Nothing Then
        Set GetSendWorkbookAttachmentOutlookApp = CreateObject("Outlook.Application")
    End If
End Function