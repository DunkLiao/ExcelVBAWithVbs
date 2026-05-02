Attribute VB_Name = "UnzipFile"
Option Explicit
'*************************************************************************************
'專案名稱: 全權委託
'功能描述: 將郵件檔案解壓縮
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2018/1/19
'
'改版日期:
'改版備註: 2018/3/1 調整過濾郵件列印部份
'
'*************************************************************************************
'with full path
'"C:\Program Files\7-Zip\7z.exe" x -aoa D:\0118日報.rar -oD:\test_extract\*.* -r -pAa1234
'"C:\Program Files\7-Zip\7z.exe" x -aoa D:\新98-1續2-20180118.zip -oD:\test_extract\*.* -r -pAa1234

'no full path (used)
'"C:\Program Files\7-Zip\7z.exe" x -aoa D:\0118日報.rar -oD:\test_extract -r -pAa1234
'"C:\Program Files\7-Zip\7z.exe" x -aoa D:\新98-1續2-20180118.zip -oD:\test_extract -r -pAa1234
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

'處理保德信加密檔案
Function DealWithEncryFilePGIM(ByVal zipFileName As String, ByVal pwd As String, ByVal sender As String)
    Dim unzipFolder As String
    'zipFileName = "D:\test_extract\0118日報.rar"
    unzipFolder = GetFileFolder(zipFileName)
    StartExtractFileForPGIM zipFileName, pwd
    StartPrinting unzipFolder, sender
End Function

' 開始解壓縮檔案(for 保德信)
Function StartExtractFileForPGIM(ByVal inZipFileName As String, ByVal pwd As String)
    Dim zipFileName, destFolder, destFileName, resultFileName As String
    Dim pureZipFileName As String
    'pwd
    'pwd = "Aa1234"
    'ZIP檔案位置
    zipFileName = inZipFileName
    '解壓縮後的資料夾
    destFolder = GetFileFolder(zipFileName)
    '檔案名稱
    'destFileName = "\*.*"
    destFileName = ""

    '開始解壓縮
    ExtractFileToFolder zipFileName:=zipFileName, pwd:=pwd, destFolder:=destFolder, destFileName:=destFileName

End Function
'開始列印檔案
Function StartPrinting(ByVal folderName As String, ByVal sender As String)
    Dim fileList(), fileName, ReturnVal As Variant
    fileList = RetrivalFileListToArrary(folderName, 0)
    For Each fileName In fileList
        If PrintAttachment.ValidPrintAbleFile(fileName, sender) = True Then
            ReturnVal = ShellExecute(0&, "print", fileName, 0&, 0&, 0&)
        End If
    Next
End Function

'取得檔案資料夾
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
    GetFileFolder = myFile.ParentFolder
    Set myFile = Nothing
    Set fso = Nothing
End Function
'解壓縮檔案
Function ExtractFileToFolder(ByVal zipFileName As String, ByVal pwd As String, ByVal destFolder As String, ByVal destFileName As String)
    Dim zipExePath, cmdLine As String
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    '7zip執行檔位置
    zipExePath = "C:\Program Files\7-Zip\7z.exe"
    cmdLine = """" & zipExePath & """  x -aoa " & zipFileName & " -o" & destFolder & destFileName & " -r -p" & pwd
    'Debug.Print cmdLine
    wsh.Run cmdLine, windowStyle, waitOnReturn
End Function

'列出檔案清單
'depth=0
Function RetrivalFileListToArrary(ByVal strDir As String, ByRef depth As Integer)
    Dim thePath As String
    Dim strSdir As String
    Dim theDirs As Scripting.Folders
    Dim theDir As Scripting.Folder
    Dim theFile As Scripting.File
    Dim myFSO As Scripting.FileSystemObject
    Dim subFolderCount As Integer
    Dim fileList() As Variant
    Dim counter As Long

    ReDim fileList(65536)
    counter = 0

    Set myFSO = New Scripting.FileSystemObject
    If Right(strDir, 1) <> "" Then strDir = strDir & ""
    thePath = thePath & strDir

    '列出第一層根目錄的檔案
    If depth = 0 Then
        For Each theFile In myFSO.GetFolder(strDir).Files
            fileList(counter) = theFile.Path
            counter = counter + 1
        Next
        depth = 1
    End If

    '尋找所有子目錄的檔案
    Set theDirs = myFSO.GetFolder(strDir).SubFolders
    For Each theDir In theDirs
        For Each theFile In theDir.Files
            fileList(counter) = theFile.Path
            counter = counter + 1
        Next
        RetrivalFileListToArrary strDir:=theDir.Path, depth:=depth
    Next

    Set myFSO = Nothing
    If counter > 0 Then
        counter = counter - 1
    End If
    ReDim Preserve fileList(counter)

    RetrivalFileListToArrary = fileList
End Function
