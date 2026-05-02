Attribute VB_Name = "SendNews"
Option Explicit
'*************************************************************************************
'專案名稱: 金融資訊專區檔案下載
'功能描述: 寄送剪報圖片檔案
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/6/15
'
'改版日期:
'改版備註: 2017/6/16 調整收信人清單及主旨產生方式
'*************************************************************************************

'開始寄送郵件
Function StartSend()
    Dim folderName As String
    folderName = ThisWorkbook.Sheets("總表").Range("F2").Text
    SendMailWithBotNews folderName:=folderName
End Function

'取得附件清單
Function SendMailWithBotNews(ByVal folderName As String)
'記得引用Microsoft Scripting Runtime
    Dim fso As Scripting.FileSystemObject
    Dim myFile As Scripting.File
    Dim subject, attFile As String
    Dim resultSubjectList(), cCMember As Variant
    Dim startIndex As Integer

    startIndex = 1
    resultSubjectList = GetCcList()
    Set fso = New Scripting.FileSystemObject

    For Each myFile In fso.GetFolder(folderName).Files
        startIndex = 1
        For Each cCMember In resultSubjectList
            '檔案類型為png，主旨即為檔名，一封信件只夾一個附檔
            If InStr(1, myFile.Name, "png") > 0 And CStr(cCMember) <> "" Then
                subject = GetSubjectTitle(Replace(myFile.Name, ".png", ""), startIndex)
                attFile = myFile.Path

                SendMail subject:=subject, cC:=CStr(cCMember), attFile:=attFile

                startIndex = startIndex + 1
            End If
        Next
    Next

    Set fso = Nothing

End Function

Function GetCcListCombine()
    Dim lastRowNumber, i As Integer
    Dim cC As String
    cC = ""
    lastRowNumber = ThisWorkbook.Sheets("寄送名單").Range("A65536").End(xlUp).Row
    For i = 1 To lastRowNumber
        If i < lastRowNumber Then
            cC = cC & ThisWorkbook.Sheets("寄送名單").Range("A" & CStr(i)).Value & ";"
        Else
            cC = cC & ThisWorkbook.Sheets("寄送名單").Range("A" & CStr(i)).Value
        End If
    Next
    GetCcListCombine = cC
End Function
'收信人清單
Function GetCcList()
    Dim lastRowNumber, i As Integer
    Dim cC() As Variant
    lastRowNumber = ThisWorkbook.Sheets("寄送名單").Range("A65536").End(xlUp).Row
    ReDim cC(lastRowNumber)
    For i = 1 To lastRowNumber
        cC(i) = ThisWorkbook.Sheets("寄送名單").Range("A" & CStr(i)).Text
    Next
    GetCcList = cC
End Function
'寄送信件函式
Function SendMail(ByVal subject As String, ByVal cC As String, ByVal attFile As String)
    Dim wbStr As String
    Dim OutlookApp As Outlook.Application
    Dim newMail As Outlook.MailItem
    Dim myAttachments As Outlook.Attachments
    Set OutlookApp = New Outlook.Application
    Set newMail = OutlookApp.CreateItem(olMailItem)
    With newMail
        .subject = subject
        .To = cC
        Set myAttachments = newMail.Attachments
        myAttachments.Add attFile
        .Send
    End With
    Set newMail = Nothing
    Set myAttachments = Nothing
    Set OutlookApp = Nothing
End Function
'取得主旨名稱
Function GetSubjectTitle(ByVal fileName As String, ByVal index As Integer)
    Dim subject As String
    '前字元+檔名+後字元
    subject = ThisWorkbook.Sheets("寄送名單").Range("B" & CStr(index)).Text & fileName _
              & ThisWorkbook.Sheets("寄送名單").Range("C" & CStr(index)).Text
    GetSubjectTitle = subject
End Function


