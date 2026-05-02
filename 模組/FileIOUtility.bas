Attribute VB_Name = "FileIOUtility"
''檔案工具模組
''作者：Guan Jhih Liao
''2017/4/6 增加寫入CSV功能

'寫入檔案
Sub WriteFile(ByVal fileName As String, ByVal content As String)
    Open fileName For Output As #1
    Print #1, content
    Close #1
End Sub

Sub WriteFileByDefaultEncoding(ByVal fileName As String, ByVal content As String)
'ensure reference is set to Microsoft ActiveX DataObjects library (the latest version of it).
'under "tools/references"... references travel with the excel file, so once added, no need to worry.
'if not you will get a type mismatch / library error on line below.
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2    'Specify stream type - we want To save text/string data.
    fsT.Charset = "_autodetect"    'Specify charset For the source text data.
    fsT.Open    'Open the stream And write binary data To the object
    fsT.WriteText content
    fsT.SaveToFile fileName, 2    'Save binary data To disk
    Set fsT = Nothing
End Sub
'使用檔案總管開啟檔案或資料夾
Function OpenFileWithExplore(ByVal fileName As String)
'Microsoft Shell Controls And Automation設定引用項目
    Dim mySh As Shell32.Shell
    Set mySh = CreateObject("Shell.Application")
    mySh.Open fileName      '任意的資料夾或檔案
    Set mySh = Nothing              '釋放物件
End Function

'讀取檔案
'記得設定Microsoft Scripting Runtime設定引用項目
Function ReadFile(ByVal infilename As String)

    Dim myFSO As Scripting.FileSystemObject
    Dim myTxt As Scripting.TextStream
    Dim myStr As String
    Dim resultString() As Variant
    ReDim resultString(65536)
    Dim rowNumber As Integer
    rowNumber = 0

    Set myFSO = CreateObject("Scripting.FileSystemObject")
    '指定檔案名稱
    Set myTxt = myFSO.OpenTextFile(fileName:=infilename, _
                                   IOMode:=ForReading)
    With myTxt
        Do Until .AtEndOfStream
            resultString(rowNumber) = CStr(.ReadLine)
            rowNumber = rowNumber + 1
        Loop
        .Close
    End With
    'Debug.Print rowNumber
    ReDim Preserve resultString(rowNumber)
    Set myTxt = Nothing                            '釋放物件
    Set myFSO = Nothing
    ReadFile = resultString

End Function

'列出檔案清單
'depth=0
Function RetrivalFileList(ByVal strDir As String, ByRef myRange As Range, ByRef depth As Integer)
    Dim thePath As String
    Dim strSdir As String
    Dim theDirs As Scripting.Folders
    Dim theDir As Scripting.Folder
    Dim theFile As Scripting.File
    Dim myFSO As Scripting.FileSystemObject
    Dim subFolderCount As Integer

    Set myFSO = New Scripting.FileSystemObject
    If Right(strDir, 1) <> "" Then strDir = strDir & ""
    thePath = thePath & strDir

    '列出第一層根目錄的檔案
    If depth = 0 Then
        For Each theFile In myFSO.getfolder(strDir).Files
            myRange = theFile.path
            myRange.Next = theFile.Size
            myRange.Next.Next = theFile.DateLastModified
            Set myRange = myRange.Offset(1, 0)
        Next
        depth = 1
    End If

    '尋找所有子目錄的檔案
    Set theDirs = myFSO.getfolder(strDir).SubFolders
    For Each theDir In theDirs
        For Each theFile In theDir.Files
            myRange = theFile.path
            myRange.Next = theFile.Size
            myRange.Next.Next = theFile.DateLastModified
            Set myRange = myRange.Offset(1, 0)
        Next
        RetrivalFileList strDir:=theDir.path, myRange:=myRange, depth:=depth
    Next
    Set myFSO = Nothing
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
        For Each theFile In myFSO.getfolder(strDir).Files
            fileList(counter) = theFile.path
            counter = counter + 1
        Next
        depth = 1
    End If

    '尋找所有子目錄的檔案
    Set theDirs = myFSO.getfolder(strDir).SubFolders
    For Each theDir In theDirs
        For Each theFile In theDir.Files
            fileList(counter) = theFile.path
            counter = counter + 1
        Next
        RetrivalFileListToArrary strDir:=theDir.path, depth:=depth
    Next

    Set myFSO = Nothing
    If counter > 0 Then
        counter = counter - 1
    End If
    ReDim Preserve fileList(counter)

    RetrivalFileListToArrary = fileList
End Function


'列出所有子目錄名稱大小及最後修改日期
Function RetrivalAllSubFolderList(ByVal strDir As String, ByRef myRange As Range)
    Dim thePath As String
    Dim strSdir As String
    Dim theDirs As Scripting.Folders
    Dim theDir As Scripting.Folder
    Dim theFile As Scripting.File
    Dim myFSO As Scripting.FileSystemObject
    Dim subFolderCount As Integer

    Set myFSO = New Scripting.FileSystemObject
    If Right(strDir, 1) <> "" Then strDir = strDir & ""
    thePath = thePath & strDir

    '尋找所有子目錄
    Set theDirs = myFSO.getfolder(strDir).SubFolders
    For Each theDir In theDirs
        myRange = theDir.path
        myRange.Next = theDir.Size
        myRange.Next.Next = theDir.DateLastModified
        Set myRange = myRange.Offset(1, 0)
        RetrivalAllSubFolderList strDir:=theDir.path, myRange:=myRange
    Next
    Set myFSO = Nothing
End Function

'存檔並取得檔名
Function SaveAs(initialName As String) As String
    On Error GoTo EndNow
    With Application.FileDialog(msoFileDialogSaveAs)
        .AllowMultiSelect = False
        .ButtonName = "&Save As"
        .InitialFileName = initialName
        .Title = "File Save As"
        .Show
        SaveAs = .SelectedItems(1)
    End With
EndNow:
End Function

'取得磁碟機類型
'用法
'   strDriveType = GetDriveType(objDrive.drivetype)
'   If Not objDrive.IsReady Then
'      MsgBox strDriveType + "[ " + Mid(ComboBox2.Text, 1, 1) + " ]尚未備妥，無法轉換", vbOKOnly + vbExclamation, "提示訊息"
'      Exit Function
'   End If
Public Function GetDriveType(ByVal nType As Long) As String
    Select Case nType
    Case Unknown
        GetDriveType = "無從判斷"
    Case Removable
        GetDriveType = "磁碟機"
    Case Fixed
        GetDriveType = "硬碟"
    Case Remote
        GetDriveType = "網路磁碟機"
    Case CDRom
        GetDriveType = "光碟機"
    Case RamDisk
        GetDriveType = "RamDisk"
    End Select
End Function

'檢查磁碟機是否存在
'用法
'If ChkDrive(strFile) Then
'      If Dir(strFile) = "" Then
'         MsgBox "找不到資料檔" + strFile, vbOKOnly + vbExclamation, "提示訊息"
'         Exit Sub
'      End If
'   Else
'      MsgBox "磁碟機未備妥", vbOKOnly + vbExclamation, "提示訊息"
'      Exit Sub
'   End If
Public Function ChkDrive(ByVal drive As String) As Boolean
    drive = Left(drive, 1) + ":\"
    On Error GoTo nodisk:
    Dir drive
    ChkDrive = True
    Exit Function

nodisk:
    If Err.Number = "52" Then
        ChkDrive = False
        Exit Function
    End If
End Function

'存檔並取得檔名
'用法
'    filenameKeyIn = FileIOUtility.SaveAsOtherType(tableName)
'    FileIOUtility.WriteFile filename:=filenameKeyIn, content:=createSql
'    MsgBox "建立完成!請到 " & filenameKeyIn & " 取檔"
Function SaveAsOtherType(initialName As String) As String
    On Error Resume Next
    Dim FileSelected As String

    FileSelected = Application.GetSaveAsFilename(InitialFileName:=initialName, _
                                                 FileFilter:="sql Files (*.sql), *.sql", _
                                                 Title:="Save SQL as")
    If Not FileSelected <> "False" Then
        Exit Function
    End If
    If FileSelected <> "False" Then
        SaveAsOtherType = FileSelected
        Exit Function
    End If
End Function

'把檔案讀取到workbook
'用法
'讀取來源檔案
'   On Error GoTo NotFindFile
'   LoadTxt filename:=Application.GetOpenFilename("所有檔案,*.*", , "開啟交易明細規格(長度：130)"), TargetSheet:=Sheets("來源格式檔案")
Sub LoadTxt(fileName As String, TargetSheet As Worksheet)
    Dim arrStr() As String, InputStr As String
    Dim Fn As Variant
    Dim i As Integer

    Fn = FreeFile

    Open fileName For Input As #Fn    '開啟檔案
    Application.ScreenUpdating = False    '畫面暫停更新
    i = 1
    While Not EOF(Fn)
        Line Input #Fn, InputStr    '從檔案讀出一列,
        If Len(InputStr) > 0 Then    '略過無字串的空行

            TargetSheet.Cells(i, 1) = InputStr    '把字串存到儲存格

        End If
        i = i + 1
    Wend
    Application.ScreenUpdating = True    '畫面恢復更新
    Close #Fn

End Sub

'寫入CSV檔案(使用CurrentRegion)
Function WriteToCsv(ByVal fileName As String, ByVal sheetName As String, ByVal x As Integer, ByVal y As Integer)
'Microsoft Forms 2.0 Object Library 設定引用項目(C:\WINDOWS\system32\FM20.DLL)
'Microsoft Scripting Runtime 設定引用項目
    Dim myFileName As String
    Dim myDataobj As DataObject
    Dim myFSO As Scripting.FileSystemObject
    Dim myTst As Scripting.TextStream
    Dim myStr As String
    Dim myDataRng As Range
    myFileName = fileName                  '儲存檔案名
    On Error Resume Next
    Kill myFileName     '刪除同名的檔案
    On Error GoTo 0
    '指定儲存的表
    Set myDataRng = Worksheets(sheetName).Cells(y, x).CurrentRegion
    '將資料送到剪貼簿
    myDataRng.Copy
    '從剪貼簿送到字串變數
    Set myDataobj = New DataObject
    myDataobj.GetFromClipboard
    myStr = myStr & Replace(myDataobj.GetText, vbTab, ",")
    Set myDataobj = Nothing                       '釋放物件
    '從字串變數送到檔案
    Set myFSO = New Scripting.FileSystemObject
    Set myTst = myFSO.OpenTextFile( _
                fileName:=myFileName, _
                IOMode:=ForWriting, Create:=True)
    myTst.Write myStr
    myTst.Close
    Set myTst = Nothing                           '釋放物件
    Set myFSO = Nothing
End Function

'寫入CSV檔案(使用指定的Range)
Function WriteToCsvWithSelectRange(ByVal fileName As String, ByVal sheetName As String, ByVal selectRange As Range)
'Microsoft Forms 2.0 Object Library 設定引用項目
'Microsoft Scripting Runtime 設定引用項目
    Dim myFileName As String
    Dim myDataobj As DataObject
    Dim myFSO As Scripting.FileSystemObject
    Dim myTst As Scripting.TextStream
    Dim myStr As String
    Dim myDataRng As Range
    myFileName = fileName                  '儲存檔案名
    On Error Resume Next
    Kill myFileName     '刪除同名的檔案
    On Error GoTo 0
    '指定儲存的表
    Set myDataRng = selectRange
    '將資料送到剪貼簿
    myDataRng.Copy
    '從剪貼簿送到字串變數
    Set myDataobj = New DataObject
    myDataobj.GetFromClipboard
    myStr = myStr & Replace(myDataobj.GetText, vbTab, ",")
    Set myDataobj = Nothing                       '釋放物件
    '從字串變數送到檔案
    Set myFSO = New Scripting.FileSystemObject
    Set myTst = myFSO.OpenTextFile( _
                fileName:=myFileName, _
                IOMode:=ForWriting, Create:=True)
    myTst.Write myStr
    myTst.Close
    Set myTst = Nothing                           '釋放物件
    Set myFSO = Nothing
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
    Set fso = Nothing
End Function

'判斷檔案是否存在
Function FileExist(ByVal fileName As String)
    Dim existFile As Boolean
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    existFile = fso.FileExists(fileName)
    Set fso = Nothing
    FileExist = existFile
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

'取得檔案名稱(不含資料夾路徑及副檔名)
Function GetFileNameWithoutFolderAndExt(ByVal fileName As String)
    Dim fName As String
    Dim existFile As Boolean
    Dim fso As Scripting.FileSystemObject
    Dim myFile As File

    Set fso = New Scripting.FileSystemObject
    existFile = fso.FileExists(fileName)
    If existFile = False Then
        GetFileNameWithoutFolderAndExt = ""
        Set myFile = Nothing
        Set fso = Nothing
        Exit Function
    End If

    Set myFile = fso.GetFile(fileName)
    GetFileNameWithoutFolderAndExt = fso.GetBaseName(fileName)
    Set myFile = Nothing
    Set fso = Nothing
End Function

'取得檔案副檔名
Function GetFileNameExtension(ByVal fileName As String)
    Dim fName As String
    Dim existFile As Boolean
    Dim fso As Scripting.FileSystemObject

    Set fso = New Scripting.FileSystemObject
    existFile = fso.FileExists(fileName)
    If existFile = False Then
        GetFileNameExtension = ""
        Set fso = Nothing
        Exit Function
    End If

    GetFileNameExtension = fso.GetExtensionName(fileName)
    Set fso = Nothing
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

'刪除資料夾下所有的檔案
Function DelFolder(ByVal folderName As String)
    Call Shell("cmd.exe /C DEL /S /Q """ & folderName & """", vbHide)
End Function

'複製資料夾結構
Function CopyFolderStructure(ByVal sourcefolderName As String, ByVal destFolderName As String)
    Call Shell("cmd.exe /c xcopy /T /E """ & sourcefolderName & """ """ & destFolderName & """", vbHide)
End Function

'開啟資料夾
Function OpenFolder(ByVal folderName As String)
    Call Shell("cmd.exe /c start """" """ & folderName & """", vbHide)
End Function

'讀取檔案
'記得設定Microsoft Scripting Runtime設定引用項目
Function ReadFileToString(ByVal infilename As String)

    Dim myFSO As Scripting.FileSystemObject
    Dim myTxt As Scripting.TextStream
    Dim myStr As String
    Dim resultString As String

    Set myFSO = CreateObject("Scripting.FileSystemObject")
    '指定檔案名稱
    Set myTxt = myFSO.OpenTextFile(fileName:=infilename, _
                                   IOMode:=ForReading)
    With myTxt
        Do Until .AtEndOfStream
            resultString = resultString + CStr(.ReadLine)
        Loop
        .Close
    End With

    Set myTxt = Nothing                            '釋放物件
    Set myFSO = Nothing
    ReadFileToString = resultString

End Function
