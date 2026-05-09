Attribute VB_Name = "ValidateAndCleanEmail"
Option Explicit
'*************************************************************************************
'模組名稱: ValidateAndCleanEmail
'功能說明: 驗證並清理 Email 欄位，統一轉小寫、移除多餘空白，並標記格式不符的資料
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

Sub TestValidateAndCleanEmail()
    Dim ws As Worksheet
    Set ws = GetOrCreateEmailSheet(ThisWorkbook, "Email清理範例")
    Call FillDirtyEmailData(ws)
    Call CleanAndValidateEmailColumn(ws, 2, 3)
    ws.Columns("A:C").AutoFit
    MsgBox "Email 驗證與清理完成！", vbInformation, "完成"
End Sub

' 清理並驗證 Email 欄位，結果寫入狀態欄
Sub CleanAndValidateEmailColumn(ByVal ws As Worksheet, ByVal emailCol As Long, ByVal statusCol As Long)
    Dim lastRow  As Long
    Dim r        As Long
    Dim strVal   As String

    Application.ScreenUpdating = False
    ws.Cells(1, statusCol).Value = "驗證狀態"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        strVal = Trim(LCase(CStr(ws.Cells(r, emailCol).Value)))
        ' 移除全形空格
        strVal = Replace(strVal, Chr(12288), "")
        ws.Cells(r, emailCol).Value = strVal

        ' 基本格式驗證：包含 @ 且 @ 後有 .
        If IsValidEmailFormat(strVal) Then
            ws.Cells(r, statusCol).Value = "格式正確"
            ws.Cells(r, statusCol).Interior.Color = RGB(198, 239, 206)
        Else
            ws.Cells(r, statusCol).Value = "格式錯誤"
            ws.Cells(r, statusCol).Interior.Color = RGB(255, 199, 206)
        End If
    Next r

    Application.ScreenUpdating = True
End Sub

' 基本 Email 格式驗證（含 @，且 @ 後段含 .）
Private Function IsValidEmailFormat(ByVal emailStr As String) As Boolean
    Dim atPos  As Integer
    Dim dotPos As Integer

    atPos = InStr(emailStr, "@")
    If atPos < 2 Then
        IsValidEmailFormat = False
        Exit Function
    End If
    dotPos = InStr(atPos + 1, emailStr, ".")
    If dotPos = 0 Or dotPos = Len(emailStr) Then
        IsValidEmailFormat = False
        Exit Function
    End If
    IsValidEmailFormat = True
End Function

Private Sub FillDirtyEmailData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("姓名", "Email")
    ws.Range("A2:B2").Value = Array("王大明", " WANG@EXAMPLE.COM ")
    ws.Range("A3:B3").Value = Array("李小花", "Lee@Test.com.tw")
    ws.Range("A4:B4").Value = Array("陳志強", "chen.email.com")
    ws.Range("A5:B5").Value = Array("林美雪", " LIN@Company.ORG ")
    ws.Range("A6:B6").Value = Array("張建國", "chang@")
    ws.Columns("A:B").AutoFit
End Sub

Private Function GetOrCreateEmailSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateEmailSheet = ws
End Function