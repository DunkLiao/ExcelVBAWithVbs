Attribute VB_Name = "FileSave"
Option Explicit
'*************************************************************************************
'專案名稱: 雪梨分行APRA報表
'功能描述: 將報表結果另存新檔
'
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2015/9/4
'
'改版日期:
'改版備註:
'*************************************************************************************

'選取所有的工作頁，並另存新檔
Public Function SaveAllSheetToNewFile(ByVal inFileName As String)

Dim sheetNames() As String
sheetNames = GelAllSheetName()

Sheets(sheetNames).Select
Sheets(sheetNames).Copy

Application.DisplayAlerts = False
ActiveWorkbook.SaveAs filename:=inFileName, _
        FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
ActiveWorkbook.Close
Application.DisplayAlerts = True

End Function


'取得所有的工作頁名稱
Public Function GelAllSheetName() As String()

Dim resultString() As String
ReDim resultString(Sheets.Count)
Dim i As Long
For i = 1 To Sheets.Count
 resultString(i - 1) = Sheets(i).Name
Next
ReDim Preserve resultString(Sheets.Count - 1)
GelAllSheetName = resultString

End Function

