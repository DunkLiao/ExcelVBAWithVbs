Attribute VB_Name = "RemoveMail"
Option Explicit
'*************************************************************************************
'專案名稱: 期信基金帳務處理
'功能描述:
' 將基金收件匣未讀取信件的轉寄email刪除

'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/11/8
'
'改版日期:
'改版備註:
'
'
'*************************************************************************************

'將基金收件匣未讀取信件的轉寄email刪除
Sub A_Step9_RemoveFundEmail()
    Dim myNameSpace As Outlook.NameSpace
    Dim myFolder As Outlook.MAPIFolder
    Dim myAttachments As Outlook.Attachments
    Dim myItems As Outlook.Items
    Dim TargetFolder As String, SFName As String, NSFName As String
    Dim resultFolder As String
    Dim i As Integer
    Dim fs As Variant
    Dim mail, att, ReturnVal As Variant

    '重建資料夾，執行前先清除有可能造成鎖定的檔案
    TargetFolder = "D:\來信的附件檔"
    A_Step2_CloseImgFiles
    Sleep (2000)
    MakeDir (TargetFolder)

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.Folders("個人資料夾").Folders("01.工作").Folders("01.基金")
    'myFolder 代表 "收件匣 Inbox"
    Set myItems = myFolder.Items
    'myItems 代表 "收件匣" 中所有信件 (的集合)
    For Each mail In myItems   '檢查每一封信
        If mail.UnRead = True Then
            If InStr(1, CStr(mail.Subject), "FW: ") > 0 Then
                mail.Delete
            End If
        End If
    Next mail
End Sub

