Attribute VB_Name = "GenReportA"
Option Explicit
'*************************************************************************************
'專案名稱: 雪梨分行APRA報表
'功能描述: 執行報表功能(產生年報)
'
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2015/9/4
'
'改版日期:
'改版備註:
'*************************************************************************************

'產生報表
Public Function GenReport(ByVal dataDate As String) As String
 '設定資料日期
 On Error GoTo ERRORHANDLE1
 Call MenuGenMrpt.SetDataDate(dataDate)
 
 '重新設定畫面
 On Error GoTo ERRORHANDLE2
 Call MenuGenMrpt.CmdReset_Click
 
 '讀取報表資料
 On Error GoTo ERRORHANDLE3
 Call MenuGenMrpt.CmdGenMP_Click
 
 '針對金額欄位進行格式調整
 On Error GoTo ERRORHANDLE4
 Call MenuGenMrpt.CmdRound_Click
   
 '關閉資料庫連線及畫面
 On Error GoTo ERRORHANDLE5
 Call MenuGenMrpt.CmdExit_Click
 GoTo ALLOK

ERRORHANDLE1:
    GenReport = "設定資料日期發生錯誤!"
ERRORHANDLE2:
    GenReport = "重新設定畫面發生錯誤!"
ERRORHANDLE3:
    GenReport = "讀取報表資料發生錯誤!"
ERRORHANDLE4:
    GenReport = "針對金額欄位進行格式調整發生錯誤!"
ERRORHANDLE5:
    GenReport = "關閉資料庫連線及畫面發生錯誤!"
ALLOK:
    GenReport = "Finish!"
End Function

'將產生結果另存新檔
Public Function SaveReportToNewFile(ByVal filename As String)

'另存新檔
Application.DisplayAlerts = False
 On Error GoTo ERRORHANDLE6
 Call FileSave.SaveAllSheetToNewFile(filename)
Application.DisplayAlerts = True
 
GoTo ALLOK
 
ERRORHANDLE6:
    SaveReportToNewFile = "另存新檔發生錯誤!"
ALLOK:
    SaveReportToNewFile = "Finish!"
End Function
