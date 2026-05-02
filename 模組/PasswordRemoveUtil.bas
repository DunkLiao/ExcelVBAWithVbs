Attribute VB_Name = "PasswordRemoveUtil"
Option Explicit
'*************************************************************************************
'專案名稱: 全委帳務處理
'功能描述: 將excel檔案的密碼移除
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/7/25
'
'改版日期:
'改版備註:
'
'*************************************************************************************

'將原有密碼移除後，以原檔名存檔
Function RemovePassord(ByVal fileName As String, ByVal pwd As String)
    Workbooks.Open fileName:= _
                   fileName, Password:=pwd
    ActiveWorkbook.SaveAs fileName:= _
                          fileName, _
                          FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
                          ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWindow.Close
End Function

