Attribute VB_Name = "SendMail"
Option Explicit
'*************************************************************************************
'專案名稱: 全委帳務處理
'功能描述: 寄送股款信件
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期：2017/6/13
'
'改版日期:
'改版備註:
'
'*************************************************************************************
Sub B_Step1_SendMailToStock()
    Dim tradeDay, sender As String
    tradeDay = ""
    tradeDay = InputBox("請輸入股款年月日", "股款設定")
    If tradeDay = "" Then
        MsgBox "請輸入股款年月日!"
        Exit Sub
    End If

    sender = "abc@def.com"

    SendMailToStock tradeDay:=tradeDay, sender:=sender
End Sub

'寄送股款信件
Function SendMailToStock(ByVal tradeDay As String, ByVal sender As String)
    Dim wbStr As String
    Dim OutlookApp As Outlook.Application
    Dim newMail As Outlook.MailItem
    Dim myAttachments As Outlook.Attachments
    Set OutlookApp = New Outlook.Application
    Set newMail = OutlookApp.CreateItem(olMailItem)
    With newMail
        .Subject = tradeDay & "股款"
        .To = sender
        Set myAttachments = newMail.Attachments
        myAttachments.Add "D:\" & tradeDay & "取款.zip"
        myAttachments.Add "D:\" & tradeDay & "匯款.zip"
        .Send
    End With
    MsgBox "郵件發送成功!"
    Set newMail = Nothing
    Set myAttachments = Nothing
    Set OutlookApp = Nothing
End Function
