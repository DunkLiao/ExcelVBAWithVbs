Attribute VB_Name = "SendingMail"
Option Explicit
'*************************************************************************************
'專案名稱: 底層元件
'功能描述: 寄送email
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/9/29
'
'改版日期:
'改版備註:
'
'*************************************************************************************
'Sub test()
'    Dim list, bccList As String
'    list = "killysss.ec781f@m.evernote.com"
'    bccList = "134719@mail.bot.com.tw"
'    SendMailWithBcc subject:="test", ccList:=list, bccList:=bccList, ToList:=list, htmlBody:="好"
'End Sub

'寄送信件函式
'引用 Microsoft Outlook 12.0 object library
Function SendMail(ByVal subject As String, ByVal ccList As String, ByVal ToList As String, ByVal htmlBody As String)
    Dim wbStr As String
    Dim OutlookApp As Outlook.Application
    Dim newMail As Outlook.MailItem
    Dim myAttachments As Outlook.Attachments
    Set OutlookApp = New Outlook.Application
    Set newMail = OutlookApp.CreateItem(olMailItem)
    With newMail
        .subject = subject
        .BodyFormat = olFormatHTML
        .To = ToList
        .cC = ccList
        .htmlBody = htmlBody
        .Send
    End With
    Set newMail = Nothing
    Set myAttachments = Nothing
    Set OutlookApp = Nothing
End Function


'寄送信件函式(密件複本)
'引用 Microsoft Outlook 12.0 object library
Function SendMailWithBcc(ByVal subject As String, ByVal ccList As String, ByVal bccList As String, ByVal ToList As String _
                                                                                                   , ByVal htmlBody As String)
    Dim wbStr As String
    Dim OutlookApp As Outlook.Application
    Dim newMail As Outlook.MailItem
    Dim myAttachments As Outlook.Attachments
    Set OutlookApp = New Outlook.Application
    Set newMail = OutlookApp.CreateItem(olMailItem)
    With newMail
        .subject = subject
        .BodyFormat = olFormatHTML
        .To = ToList
        .cC = ccList
        .BCC = bccList
        .htmlBody = htmlBody
        .Send
    End With
    Set newMail = Nothing
    Set myAttachments = Nothing
    Set OutlookApp = Nothing
End Function
