VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'*************************************************************************************
'專案名稱: 全委帳務處理
'功能描述: 自動密件副本
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/9/29
'
'改版日期:
'改版備註:
'
'*************************************************************************************

'自動密件副本
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim emailSettingName As String
    Dim myBccList As String

    emailSettingName = "134719@mail.bot.com.tw"
    myBccList = "134719@mail.bot.com.tw"

    If Item.Class = olMail Then
        If Item.SendUsingAccount.DisplayName = emailSettingName Then
            Dim oRecipient As Recipient
            Set oRecipient = Item.Recipients.Add(myBccList)
            oRecipient.Type = olBCC
            oRecipient.Resolve
            Set oRecipient = Nothing
        End If
    End If

End Sub
