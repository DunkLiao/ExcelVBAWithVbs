Attribute VB_Name = "PrintAttachment"
Option Explicit
'*************************************************************************************
'專案名稱: 全委帳務處理
'功能描述:
' email 收件自動下載列印

'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/6/8
'
'改版日期:
'改版備註: 2017/6/13 增加關閉Excel,RAR及adobe檔案功能(因為會與重建資料夾功能衝突，先關閉後再執行)
'                 2017/6/14 調整關閉Excel,RAR及adobe檔案功能執行順序及執行不顯示視窗
'                 2017/6/26 調整排除建立資料夾時的特殊字元
'                 2017/12/27 增加全委代理設定
'                 2018/1/19 增加保德信檔案處理
'                 2018/2/7 增加儲存基金對帳單檔案及安聯檔案
'                 2018/2/13 調整代理部份
'                 2018/3/1 調整過濾郵件列印部份
'                 2018/8/28 增加判斷小寫檔名
'*************************************************************************************



#If VBA7 Then
    Declare PtrSafe  Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)    'For 64 Bit Systems
    Declare PtrSafe  Function ShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
                                   ByVal lpFile As String, ByVal lpParameters As String, _
                                   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)    'For 32 Bit Systems
    Declare Function ShellExecute Lib "shell32.dll" _
                                  Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
                                                         ByVal lpFile As String, ByVal lpParameters As String, _
                                                         ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Sub A_測試()
    Sleep (3000)
    MsgBox ("!!")
End Sub

'將收件匣未讀取信件的所有附件檔另存
Sub A_Step1_SaveAttachments()
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
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    'myFolder 代表 "收件匣 Inbox"
    Set myItems = myFolder.Items
    'myItems 代表 "收件匣" 中所有信件 (的集合)
    For Each mail In myItems   '檢查每一封信
        If mail.UnRead = True Then
            resultFolder = TargetFolder & "\" & GetFolderName(mail.Subject)
            MakeDir (resultFolder)
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
                    If checkname(NSFName) = True Then
                        att.SaveAsFile NSFName    '用加了數字的檔名儲存
                    End If
                Else
                    If checkname(SFName) = True Then
                        att.SaveAsFile SFName  ''若檔案不存在, 就用原來的檔名儲存

                        If ValidPrintAbleFile(SFName, mail.SenderEmailAddress) = True Then
                            ReturnVal = ShellExecute(0&, "print", SFName, 0&, 0&, 0&)
                        Else
                            '處理加密的zip檔案
                            DealWithZipCryFiles Subject:=mail.Subject, SFName:=SFName, _
                                                sender:=mail.SenderEmailAddress
                        End If
                    End If
                End If
            Next att
            mail.UnRead = False
        End If
    Next mail
    'att.PrintOut
    OpenFolder2 (TargetFolder)
End Sub
'處理加密的zip檔案
Function DealWithZipCryFiles(ByVal Subject As String, ByVal SFName As String, ByVal sender As String)
'處理保德信
    If (InStr(1, Subject, "保德信") > 0 Or InStr(1, sender, "pru.com.tw") > 0) _
       And (InStr(1, SFName, ".zip") > 0 Or InStr(1, SFName, ".rar") > 0) Then
        If InStr(1, Subject, "103-1") > 0 Then
            UnzipFile.DealWithEncryFilePGIM SFName, "Aa1234", sender
        Else
            UnzipFile.DealWithEncryFilePGIM SFName, "Aa1234", sender
        End If
    End If
    '處理安聯
    If (InStr(1, Subject, "安聯") > 0 Or InStr(1, sender, "allianzgi.com") > 0) _
       And (InStr(1, SFName, ".zip") > 0 Or InStr(1, SFName, ".rar") > 0) Then
        If InStr(1, Subject, "報表") > 0 Then
            '檔案
            UnzipFile.DealWithEncryFilePGIM SFName, "A0036", sender
        Else
            '檔案
            UnzipFile.DealWithEncryFilePGIM SFName, "Allianz18", sender
        End If
        If InStr(1, sender, "Kim.Wang@allianzgi.com") > 0 Then
            '檔案
            UnzipFile.DealWithEncryFilePGIM SFName, "kimw9516", sender
        End If
    End If
End Function

Sub A_Step2_代理秉志()
    A_Step1_SavePlaceAttachments ("01.秉志")
End Sub
Sub A_Step2_代理蘭錦()
    A_Step1_SavePlaceAttachments ("02.蘭錦")
End Sub
Sub A_Step2_代理振富()
    A_Step1_SavePlaceAttachments ("03.振富")
End Sub


'將全委代理收件匣未讀取信件的所有附件檔另存
Function A_Step1_SavePlaceAttachments(ByVal userName As String)
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
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox).Folders("代理").Folders(userName)
    'myFolder 代表 "收件匣 Inbox"
    Set myItems = myFolder.Items
    'myItems 代表 "收件匣" 中所有信件 (的集合)
    For Each mail In myItems   '檢查每一封信
        If mail.UnRead = True Then
            resultFolder = TargetFolder & "\" & GetFolderName(mail.Subject)
            MakeDir (resultFolder)
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
                    If checkname(NSFName) = True Then
                        att.SaveAsFile NSFName    '用加了數字的檔名儲存
                    End If
                Else
                    If checkname(SFName) = True Then
                        att.SaveAsFile SFName  ''若檔案不存在, 就用原來的檔名儲存
                        If ValidPrintAbleFile(SFName, mail.SenderEmailAddress) = True Then
                            ReturnVal = ShellExecute(0&, "print", SFName, 0&, 0&, 0&)
                        Else
                            '處理加密的zip檔案
                            DealWithZipCryFiles Subject:=mail.Subject, SFName:=SFName, _
                                                sender:=mail.SenderEmailAddress
                        End If
                    End If
                End If
            Next att
            mail.UnRead = False
        End If
    Next mail
    'att.PrintOut
    OpenFolder2 (TargetFolder)
End Function

'將基金收件匣未讀取信件的所有附件檔另存
Sub A_Step1_SaveFundAttachments()
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
            resultFolder = TargetFolder & "\" & GetFolderName(mail.Subject)
            MakeDir (resultFolder)
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
                    If checkname(NSFName) = True Then
                        att.SaveAsFile NSFName    '用加了數字的檔名儲存
                    End If
                Else
                    If checkname(SFName) = True Then
                        att.SaveAsFile SFName  ''若檔案不存在, 就用原來的檔名儲存
                        If ValidPrintAbleFileFund(SFName) = True Then
                            '儲存基金對帳單檔案
                            SaveFundFiles.SaveFundFile (SFName)
                            '取消列印
                            'ReturnVal = ShellExecute(0&, "print", SFName, 0&, 0&, 0&)
                        End If
                    End If
                End If
            Next att
            mail.UnRead = False
        End If
    Next mail
    'att.PrintOut
    OpenFolder2 (TargetFolder)
End Sub

'將基金代理收件匣未讀取信件的所有附件檔另存
Sub A_Step1_SaveFundRepMgrAttachments()
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
    Set myFolder = myNameSpace.Folders("個人資料夾").Folders("01.工作").Folders("01.基金代理")
    'myFolder 代表 "收件匣 Inbox"
    Set myItems = myFolder.Items
    'myItems 代表 "收件匣" 中所有信件 (的集合)
    For Each mail In myItems   '檢查每一封信
        If mail.UnRead = True Then
            resultFolder = TargetFolder & "\" & GetFolderName(mail.Subject)
            MakeDir (resultFolder)
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
                    If checkname(NSFName) = True Then
                        att.SaveAsFile NSFName    '用加了數字的檔名儲存
                    End If
                Else
                    If checkname(SFName) = True Then
                        att.SaveAsFile SFName  ''若檔案不存在, 就用原來的檔名儲存
                        If ValidPrintAbleFileFund(SFName) = True Then
                            ReturnVal = ShellExecute(0&, "print", SFName, 0&, 0&, 0&)
                        End If
                    End If
                End If
            Next att
            mail.UnRead = False
        End If
    Next mail
    'att.PrintOut
    OpenFolder2 (TargetFolder)
End Sub

'關閉多開的檔案
Sub A_Step2_CloseImgFiles()
'關閉多開的Excel檔案
    CallTaskKill ("EXCEL.EXE")
    '關閉多開的Rar檔案
    CallTaskKill ("WinRAR.exe")
    '關閉多開的Adobe檔案
    CallTaskKill ("Acrobat.exe")
End Sub

Function checkname(ByVal fileName As String)
    checkname = True
End Function

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
Function MakeDir(ByVal folderName As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    If fso.FolderExists(folderName) = True Then
        fso.DeleteFolder folderName, force:=True
        fso.CreateFolder folderName
    Else
        fso.CreateFolder folderName
    End If
End Function

'區分可以列印的檔案名稱與種類
Function ValidPrintAbleFile(ByVal fileName As String, ByVal sender As String) '
' for 保德信
    If InStr(1, sender, "pru.com.tw") > 0 Then
        If InStr(1, fileName, "短期投資未到期明細表") > 0 Then
            ValidPrintAbleFile = False
            Exit Function
        End If
        If InStr(1, fileName, "資產投資組合表") > 0 Then
            ValidPrintAbleFile = False
            Exit Function
        End If
    End If
    'for 安聯
    If InStr(1, sender, "allianzgi.com") > 0 Then
        If InStr(1, fileName, "短期票券應收利息明細表") > 0 Then
            ValidPrintAbleFile = False
            Exit Function
        End If
    End If
    'for 台新
    If InStr(1, sender, "tsit.com.tw") > 0 Then
        If InStr(1, fileName, "附買回債券票券應計利息明細表") > 0 Then
            ValidPrintAbleFile = False
            Exit Function
        End If
    End If
    Dim fileNameLow As String
    '判斷小寫檔名
    fileNameLow = LCase(fileName)
    If InStr(1, fileNameLow, "doc") > 0 Or InStr(1, fileNameLow, "docx") > 0 Or InStr(1, fileNameLow, "rtf") > 0 _
       Or InStr(1, fileNameLow, "pdf") > 0 Then
        ValidPrintAbleFile = True
    Else
        ValidPrintAbleFile = False
    End If

End Function

'區分可以列印的檔案名稱與種類(基金檔案)
Function ValidPrintAbleFileFund(ByVal fileName As String) '
' for 元大期貨
    If InStr(1, fileName, "2449158C") > 0 Or _
       InStr(1, fileName, "2518539C") > 0 Or _
       InStr(1, fileName, "2533187C") > 0 Or _
       InStr(1, fileName, "9756162C") > 0 Then
        ValidPrintAbleFileFund = False
        Exit Function
    End If

    Dim fileNameLow As String
    '判斷小寫檔名
    fileNameLow = LCase(fileName)
    If InStr(1, fileNameLow, "doc") > 0 Or InStr(1, fileNameLow, "docx") > 0 Or InStr(1, fileNameLow, "rtf") > 0 _
       Or InStr(1, fileNameLow, "pdf") > 0 Then
        ValidPrintAbleFileFund = True
    Else
        ValidPrintAbleFileFund = False
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


Function CallTaskKill(ByVal imgName As String)
    Call Shell("cmd.exe /C ""taskkill /F /IM " & imgName & """", vbHide)
End Function

