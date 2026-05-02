Attribute VB_Name = "ODBCUtility"
Option Explicit
'*************************************************************************************
'專案名稱: 雪梨分行APRA報表
'功能描述: 讀取ODBC資料源
'http://stackoverflow.com/questions/164967/how-can-i-enumerate-the-list-of-dsns-set-up-on-a-computer-using-vba
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2016/5/30
'
'改版日期:
'改版備註:
'*************************************************************************************
  Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
      Alias "RegOpenKeyExA" _
      (ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal ulOptions As Long, _
      ByVal samDesired As Long, phkResult As Long) As Long

  Private Declare Function RegEnumValue Lib "advapi32.dll" _
      Alias "RegEnumValueA" _
      (ByVal hKey As Long, _
      ByVal dwIndex As Long, _
      ByVal lpValueName As String, _
      lpcbValueName As Long, _
      ByVal lpReserved As Long, _
      lpType As Long, _
      lpData As Any, _
      lpcbData As Long) As Long

  Private Declare Function RegCloseKey Lib "advapi32.dll" _
      (ByVal hKey As Long) As Long

  Const HKEY_CLASSES_ROOT = &H80000000
  Const HKEY_CURRENT_USER = &H80000001
  Const HKEY_LOCAL_MACHINE = &H80000002
  Const HKEY_USERS = &H80000003

  Const ERROR_SUCCESS = 0&

  Const SYNCHRONIZE = &H100000
  Const STANDARD_RIGHTS_READ = &H20000
  Const STANDARD_RIGHTS_WRITE = &H20000
  Const STANDARD_RIGHTS_EXECUTE = &H20000
  Const STANDARD_RIGHTS_REQUIRED = &HF0000
  Const STANDARD_RIGHTS_ALL = &H1F0000
  Const KEY_QUERY_VALUE = &H1
  Const KEY_SET_VALUE = &H2
  Const KEY_CREATE_SUB_KEY = &H4
  Const KEY_ENUMERATE_SUB_KEYS = &H8
  Const KEY_NOTIFY = &H10
  Const KEY_CREATE_LINK = &H20
  Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                    KEY_QUERY_VALUE Or _
                    KEY_ENUMERATE_SUB_KEYS Or _
                    KEY_NOTIFY) And _
                    (Not SYNCHRONIZE))

  Const REG_DWORD = 4
  Const REG_BINARY = 3
  Const REG_SZ = 1
  
  '設定ODBC連線資料源
  Public Function GetConnectedODBC()
  Dim odbcs As String
  Dim resultODBC As String
  Dim arrStr() As String
  Dim j As Integer
  Dim pos As Integer
  odbcs = GetSystemDSN()
  
  If Len(odbcs) = 0 Then
    GetConnectedODBC = ""
    Exit Function
  End If
  
  If Len(odbcs) > 0 Then
        arrStr = Split(odbcs, "!")
  End If
  
  For j = 0 To UBound(arrStr)
  pos = InStr(arrStr(j), "ORSAUSY")
    If pos > 0 Then
        GetConnectedODBC = arrStr(j)
        Exit Function
    End If
  Next j
  GetConnectedODBC = ""
  End Function
  
  '列出USER DSN ODBC
  Public Function GetUserDSN()
     Dim lngKeyHandle As Long
     Dim lngResult As Long
     Dim lngCurIdx As Long
     Dim strValue As String
     Dim lngValueLen As Long
     Dim lngData As Long
     Dim lngDataLen As Long
     Dim strResult As String

     lngResult = RegOpenKeyEx(HKEY_CURRENT_USER, _
             "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", _
              0&, _
              KEY_READ, _
              lngKeyHandle)

     If lngResult <> ERROR_SUCCESS Then
         MsgBox "Cannot open key"
         Exit Function
     End If

     lngCurIdx = 0
     Do
        lngValueLen = 2000
        strValue = String(lngValueLen, 0)
        lngDataLen = 2000

        lngResult = RegEnumValue(lngKeyHandle, _
                                 lngCurIdx, _
                                 ByVal strValue, _
                                 lngValueLen, _
                                 0&, _
                                 REG_DWORD, _
                                 ByVal lngData, _
                                 lngDataLen)
        lngCurIdx = lngCurIdx + 1

        If lngResult = ERROR_SUCCESS Then
           'strResult = strResult & lngCurIdx & ": " & Left(strValue, lngValueLen) & vbCrLf
           strResult = strResult & Left(strValue, lngValueLen) & "!"
        End If
     Loop While lngResult = ERROR_SUCCESS
     Call RegCloseKey(lngKeyHandle)
     
     GetUserDSN = strResult
     'Call MsgBox(strResult, vbInformation)
  End Function
  '列出SYSTEM DSN ODBC
  Public Function GetSystemDSN()
     Dim lngKeyHandle As Long
     Dim lngResult As Long
     Dim lngCurIdx As Long
     Dim strValue As String
     Dim lngValueLen As Long
     Dim lngData As Long
     Dim lngDataLen As Long
     Dim strResult As String

     lngResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
             "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", _
              0&, _
              KEY_READ, _
              lngKeyHandle)

     If lngResult <> ERROR_SUCCESS Then
         MsgBox "Cannot open key"
         Exit Function
     End If

     lngCurIdx = 0
     Do
        lngValueLen = 2000
        strValue = String(lngValueLen, 0)
        lngDataLen = 2000

        lngResult = RegEnumValue(lngKeyHandle, _
                                 lngCurIdx, _
                                 ByVal strValue, _
                                 lngValueLen, _
                                 0&, _
                                 REG_DWORD, _
                                 ByVal lngData, _
                                 lngDataLen)
        lngCurIdx = lngCurIdx + 1

        If lngResult = ERROR_SUCCESS Then
           'strResult = strResult & lngCurIdx & ": " & Left(strValue, lngValueLen) & vbCrLf
           strResult = strResult & Left(strValue, lngValueLen) & "!"
        End If
     Loop While lngResult = ERROR_SUCCESS
     Call RegCloseKey(lngKeyHandle)
     
     GetSystemDSN = strResult
     'Call MsgBox(strResult, vbInformation)
  End Function

