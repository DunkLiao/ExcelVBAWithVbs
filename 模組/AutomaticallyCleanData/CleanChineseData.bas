Option Explicit
Attribute VB_Name = "CleanChineseData"
'*************************************************************************************
'模組名稱: CleanChineseData
'功能說明: 清理中文資料中的全形空白、全形標點符號與非預期中文字元
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestCleanChineseData()
    Call CleanChineseTextData
End Sub

' 清理中文文字資料
Sub CleanChineseTextData()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellText As String
    Dim cleanText As String

    Set ws = GetOrCreateWorksheet("中文清理範例")
    ws.Cells.Clear

    ' 建立含雜亂中文的範例資料
    ws.Range("A1").Value = "原始資料"
    ws.Range("B1").Value = "清理後"
    ws.Range("A2").Value = "　台　北　市"
    ws.Range("A3").Value = "台北市（中山區）"
    ws.Range("A4").Value = "【重要】會議通知！！！"
    ws.Range("A5").Value = "  台北市信義區  "
    ws.Range("A6").Value = "電話：02-12345678"
    ws.Range("A7").Value = "　統一編號：12345678　"

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 逐列清理
    For i = 2 To lastRow
        cellText = ws.Cells(i, 1).Value
        cleanText = CleanFullWidthSpaces(cellText)
        cleanText = CleanCommonPunctuation(cleanText)
        ws.Cells(i, 2).Value = cleanText
    Next i

    ws.Columns.AutoFit
    MsgBox "已清理 " & (lastRow - 1) & " 列中文資料。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 移除全形空白
Private Function CleanFullWidthSpaces(ByVal text As String) As String
    Dim result As String
    result = Replace(text, ChrW(12288), " ")
    CleanFullWidthSpaces = Trim(result)
End Function

' 移除常見中文全形標點
Private Function CleanCommonPunctuation(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "【", "")
    result = Replace(result, "】", "")
    CleanCommonPunctuation = result
End Function

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
