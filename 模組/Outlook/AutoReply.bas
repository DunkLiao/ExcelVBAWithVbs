Attribute VB_Name = "AutoReply"
Option Explicit
'*************************************************************************************
'專案名稱: 期信基金帳務處理
'功能描述:
' email自動回覆內容

'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/10/26
'
'改版日期:
'改版備註: 2020/3/12 增加C_Step1_台新長紅餘額OK
'
'*************************************************************************************
Sub C_Step1_ReplyReceiveFAXOK()
    ReplyAllWithSubjectAddMsg ("_已收到謝謝")
End Sub

Sub C_Step2_ReplyReceiveFAXNotOK()
    ReplyAllWithSubjectAddMsg ("_未收到請重傳FAX")
End Sub

Sub C_Step1_台新長紅餘額OK()
    ReplyWithSubjectAddMsgToUser "_餘額OK", "claire.ku@tsit.com.tw]"
End Sub

Function ReplyAllWithSubjectAddMsg(ByVal msg As String)
    Dim mail    'object/mail item iterator
    Dim replyall    'object which will represent the reply email

    For Each mail In Outlook.Application.ActiveExplorer.Selection
        If mail.Class = olMail Then
            Set replyall = mail.replyall
            With replyall
                .Subject = .Subject & msg
                .Send
            End With
        End If
    Next
End Function


Function ReplyWithSubjectAddMsgToUser(ByVal msg As String, ByVal ToList As String)
    Dim mail    'object/mail item iterator
    Dim replyall    'object which will represent the reply email

    For Each mail In Outlook.Application.ActiveExplorer.Selection
        If mail.Class = olMail Then
            Set replyall = mail.replyall
            With replyall
                .Subject = .Subject & msg
                .To = ToList
                .Send
            End With
        End If
    Next
End Function
