Attribute VB_Name = "DealOilFunAttachment"
Option Explicit
'*************************************************************************************
'專案名稱: 期信基金處理
'功能描述:
' email 收件自動下載列印(華頓石油)

'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2018/3/7
'
'改版日期:
'改版備註:
'
'*************************************************************************************


'將基金代理收件匣未讀取信件的所有附件檔另存
Sub 處理華頓石油基金()
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
    TargetFolder = "Z:\全委組帳務\帳務--新制轉檔報表\蘭錦\華頓石油"
    Sleep (2000)
    MakeDirOil (TargetFolder)

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.Folders("個人資料夾").Folders("01.工作").Folders("01.基金代理")
    'myFolder 代表 "收件匣 Inbox"
    Set myItems = myFolder.Items
    'myItems 代表 "收件匣" 中所有信件 (的集合)
    For Each mail In myItems   '檢查每一封信
        If mail.UnRead = True Then
            resultFolder = TargetFolder
            Set myAttachments = mail.Attachments
            'myAttachments 代表這封信件裡所有附件檔 (的集合)
            For Each att In myAttachments
                SFName = resultFolder & "\" & att.fileName
                If fs.FileExists(SFName) Then    '若檔案已存在, 就加上 (數字)
                    i = 0
                    Do
                        NSFName = TargetFolder & "\" & fs.GetBaseName(SFName) _
                                  & "(" & i & ")." & fs.GetExtensionName(SFName)
                        i = i + 1
                    Loop While fs.FileExists(NSFName)
                    att.SaveAsFile NSFName    '用加了數字的檔名儲存
                Else
                    att.SaveAsFile SFName  ''若檔案不存在, 就用原來的檔名儲存
                End If
            Next att
            mail.UnRead = False
        End If
    Next mail
    'att.PrintOut
    OpenFolder2 (TargetFolder)
End Sub

'排除建立資料夾時的特殊字元
Function GetFolderName(ByVal fileName As String)
    Dim strname As String
    strname = fileName
    strname = Replace(strname, "*", "_")
    strname = Replace(strname, "\", "_")
    strname = Replace(strname, "/", "_")
    strname = Replace(strname, "$", "_")
    strname = Replace(strname, "%", "_")
    strname = Replace(strname, "!", "_")
    strname = Replace(strname, "~", "_")
    strname = Replace(strname, "(", "_")
    strname = Replace(strname, ")", "_")
    strname = Replace(strname, "+", "_")
    strname = Replace(strname, ":", "_")
    strname = Replace(strname, "<", "_")
    strname = Replace(strname, ">", "_")
    strname = Replace(strname, "|", "_")
    strname = Replace(strname, "?", "_")
    strname = Replace(strname, """", "_")
    strname = Replace(strname, " ", "_")
    strname = Replace(strname, ".", "_")
    GetFolderName = strname
End Function

'建立資料夾(如果已經存在會先刪除)
Function MakeDirOil(ByVal folderName As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    If fso.FolderExists(folderName) = True Then
        fso.DeleteFolder folderName, force:=True
        fso.CreateFolder folderName
    Else
        fso.CreateFolder folderName
    End If
End Function

'開啟資料夾,另一種比較簡單的方式
Function OpenFolder2(strPath As String, Optional bnRoot As Boolean = False)
    Dim strRoot As String

    If bnRoot = True Then
        strRoot = "/root,"
    Else
        strRoot = ""
    End If

    Call Shell("explorer.exe" & " " & strRoot & strPath, vbNormalFocus)
End Function
