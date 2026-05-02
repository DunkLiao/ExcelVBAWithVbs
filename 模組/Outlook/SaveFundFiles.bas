Attribute VB_Name = "SaveFundFiles"
Option Explicit
'*************************************************************************************
'專案名稱: 期信基金
'功能描述: 儲存期貨對帳單檔案
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2018/2/7
'
'改版日期:
'改版備註:
'
'*************************************************************************************
'儲存期貨對帳單
Function SaveFundFile(ByVal sourceFileName As String)
    Dim sourceFolder, destFolder, destFileName As String
    '目的資料夾
    destFolder = "Z:\全委組帳務\帳務--新制轉檔報表\冠智\期信基金\對帳單"
    destFileName = destFolder & "\" & GetFileNameWithoutFolder(sourceFileName)
    If InStr(1, sourceFileName, ".pdf") > 0 Then
        CopyFile fileNameSource:=sourceFileName, fileNameDest:=destFileName
    End If
End Function

'複製檔案
Function CopyFile(ByVal fileNameSource As String, ByVal fileNameDest As String)
    Dim existFile As Boolean
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject


    existFile = fso.FileExists(fileNameSource)
    If existFile = False Then
        MsgBox "來源檔案不存在!"
        Set fso = Nothing
        Exit Function
    End If

    fso.CopyFile fileNameSource, fileNameDest, True
    Set fso = Nothing
End Function

'取得檔案所在資料夾
Function GetFileFolder(ByVal fileName As String)
    Dim fName As String
    Dim existFile As Boolean
    Dim fso As Scripting.FileSystemObject
    Dim myFile As File

    Set fso = New Scripting.FileSystemObject
    existFile = fso.FileExists(fileName)
    If existFile = False Then
        GetFileFolder = ""
        Set myFile = Nothing
        Set fso = Nothing
        Exit Function
    End If

    Set myFile = fso.GetFile(fileName)
    GetFileFolder = fso.GetParentFolderName(fileName)
    Set myFile = Nothing
    Set fso = Nothing
End Function

'取得檔案名稱(不含資料夾路徑)
Function GetFileNameWithoutFolder(ByVal fileName As String)
    Dim fName As String
    Dim existFile As Boolean
    Dim fso As Scripting.FileSystemObject
    Dim myFile As File

    Set fso = New Scripting.FileSystemObject
    existFile = fso.FileExists(fileName)
    If existFile = False Then
        GetFileNameWithoutFolder = ""
        Set myFile = Nothing
        Set fso = Nothing
        Exit Function
    End If

    Set myFile = fso.GetFile(fileName)
    GetFileNameWithoutFolder = fso.GetFileName(fileName)
    Set myFile = Nothing
    Set fso = Nothing
End Function
