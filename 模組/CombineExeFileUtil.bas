Attribute VB_Name = "CombineExeFileUtil"
Option Explicit
'*************************************************************************************
'專案名稱: VBA專案
'功能描述:
'與外部程式整合
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2015/11/11
'
'改版日期:
'改版備註:
'*************************************************************************************

'使用notepad，將檔案調整成unicode後，另存新檔
Public Function SaveFileToUnicode(ByVal fileName As String)
Dim R
    R = Shell("notepad.exe", 1)
    AppActivate R
    SendKeys "%FO", True '------------->[檔案][開啟檔案]
    SendKeys fileName, True  '--------->輸入檔名
    SendKeys "{ENTER}", True '-------------->按下[確定]
    SendKeys "%FA", True '------------->[檔案][另存新檔]
    SendKeys fileName, True  '--------->輸入檔名
    SendKeys "%E", True '-------------->按下[切換到編碼]
    SendKeys "unicode", True '-------------->[unicode]
    SendKeys "%S", True '-------------->按下[存檔]
    SendKeys "{TAB}", True '-------------->切換到[確定]
    SendKeys "{ENTER}", True '-------------->按下[確定]
    SendKeys "%{f4}", True '----------->按下[f4]結束程式
End Function
