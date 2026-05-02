Attribute VB_Name = "LoadConfigFile"
Option Explicit
'*************************************************************************************
'專案名稱: 雪梨分行APRA報表
'功能描述: 讀取設定檔(C:\config\Ors_AP.txt)
'
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2015/8/26
'
'改版日期:
'改版備註:
'*************************************************************************************
'帳號
Public ID As String
'密碼
Public PWD As String



Sub LoadConfig()
 Dim arrStr() As String, InputStr As String
 Dim i As Integer, j As Integer
 Dim Fn As Variant
 Dim myConfigFile As String

 Fn = FreeFile
 
 '組態檔位置
 myConfigFile = "C:\config\Ors_AP.txt"
 
 On Error GoTo Err
 Open myConfigFile For Input As #Fn
Application.ScreenUpdating = False '畫面暫停更新
i = 1: j = 1
 While Not EOF(Fn)
 Line Input #Fn, InputStr '從檔案讀出一列,
If Len(InputStr) > 0 Then '略過無字串的空行
arrStr = Split(InputStr, "=")
 '把讀入的文字列依逗號分成數個字串, 置於 arrStr 陣列裡
For j = 1 To UBound(arrStr)
    
  If arrStr(0) = "ID" Then
    ID = arrStr(1)
  End If
  If arrStr(0) = "PWD" Then
    PWD = arrStr(1)
  End If
Next j
 End If
 i = i + 1
 Wend
 Application.ScreenUpdating = True '畫面恢復更新
Close #Fn
GoTo Final

Err:
 Debug.Print "Config File Not Exist!" & myConfigFile
 
Final:
 End Sub
