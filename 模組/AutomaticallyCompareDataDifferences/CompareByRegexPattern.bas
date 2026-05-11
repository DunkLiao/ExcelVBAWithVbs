Attribute VB_Name = "CompareByRegexPattern"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByRegexPattern
'功能說明: 使用正規表示式比對工作表中指定欄位的資料格式，並標記不符合格式的儲存格
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口（預設驗證 A 欄的電子郵件格式）
Sub TestCompareByRegexPattern()
    Dim emailPattern As String
    emailPattern = "^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"
    Call ValidateColumnByRegex(ActiveSheet, 1, emailPattern, "電子郵件")
End Sub

' 使用正規表示式驗證指定欄位，標記不符合格式的儲存格
' ws           : 要驗證的工作表
' colIndex     : 要驗證的欄索引（從 1 開始）
' regexPattern : 正規表示式樣式字串
' formatDesc   : 格式描述（用於顯示訊息）
Sub ValidateColumnByRegex(ByVal ws As Worksheet, _
                           ByVal colIndex As Integer, _
                           ByVal regexPattern As String, _
                           ByVal formatDesc As String)
    On Error GoTo ErrorHandler

    Dim regex As Object
    Dim lastRow As Long
    Dim r As Long
    Dim cellVal As String
    Dim invalidCount As Long
    Dim validCount As Long

    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False
    regex.Pattern = regexPattern

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "欄位中沒有資料（第 2 列起）。", vbInformation, "無資料"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    invalidCount = 0
    validCount = 0

    ' 先清除既有的標記色
    ws.Columns(colIndex).Interior.ColorIndex = xlNone

    For r = 2 To lastRow
        cellVal = Trim(CStr(ws.Cells(r, colIndex).Value))

        If cellVal = "" Then
            ' 空白儲存格 — 標示淡黃色
            ws.Cells(r, colIndex).Interior.Color = RGB(255, 255, 180)
        ElseIf regex.Test(cellVal) Then
            ' 符合格式 — 標示淡綠色
            ws.Cells(r, colIndex).Interior.Color = RGB(198, 239, 206)
            validCount = validCount + 1
        Else
            ' 不符合格式 — 標示淡紅色
            ws.Cells(r, colIndex).Interior.Color = RGB(255, 199, 206)
            ws.Cells(r, colIndex).Font.Bold = True
            invalidCount = invalidCount + 1
        End If
    Next r

    Application.ScreenUpdating = True

    MsgBox "正規表示式驗證完成！" & vbCrLf & _
           "格式說明：" & formatDesc & vbCrLf & _
           "符合格式：" & validCount & " 筆（綠色）" & vbCrLf & _
           "不符合格式：" & invalidCount & " 筆（紅色粗體）" & vbCrLf & _
           "空白儲存格：黃色", vbInformation, "驗證結果"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "正規表示式驗證時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
