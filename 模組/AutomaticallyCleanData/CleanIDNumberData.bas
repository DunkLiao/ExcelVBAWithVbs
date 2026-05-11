Attribute VB_Name = "CleanIDNumberData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanIDNumberData
'功能說明: 清洗身分證字號欄位資料，標記格式不符或驗證失敗的記錄，並統計錯誤清單
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 字母代碼對應表（A=10, B=11, ... Z=35）
Private Function LetterCode(ByVal ch As String) As Integer
    LetterCode = Asc(UCase(ch)) - Asc("A") + 10
End Function

' 驗證台灣身分證字號格式與驗證碼
' idNum: 10 字元身分證字號
' 回傳 True 為有效
Function IsValidTWID(ByVal idNum As String) As Boolean
    Dim digits(0 To 10) As Integer
    Dim weights(0 To 10) As Integer
    Dim code As Integer
    Dim checkSum As Integer
    Dim i As Integer
    Dim ch As String

    IsValidTWID = False

    If Len(idNum) <> 10 Then Exit Function

    ch = Left(idNum, 1)
    If ch < "A" Or ch > "Z" Then Exit Function

    Dim secondCh As String
    secondCh = Mid(idNum, 2, 1)
    If secondCh <> "1" And secondCh <> "2" Then Exit Function

    Dim j As Integer
    For j = 3 To 10
        If Mid(idNum, j, 1) < "0" Or Mid(idNum, j, 1) > "9" Then Exit Function
    Next j

    ' 展開首碼字母為 2 位數
    code = LetterCode(ch)
    digits(0) = code \ 10
    digits(1) = code Mod 10

    For i = 2 To 10
        digits(i) = CInt(Mid(idNum, i, 1))
    Next i

    weights(0) = 1
    weights(1) = 9
    weights(2) = 8
    weights(3) = 7
    weights(4) = 6
    weights(5) = 5
    weights(6) = 4
    weights(7) = 3
    weights(8) = 2
    weights(9) = 1
    weights(10) = 1

    checkSum = 0
    For i = 0 To 10
        checkSum = checkSum + digits(i) * weights(i)
    Next i

    IsValidTWID = (checkSum Mod 10 = 0)
End Function

' 清洗工作表中的身分證字號欄位
' ws       : 目標工作表
' idCol    : 身分證字號欄號
' statusCol: 驗證結果輸出欄號
Sub CleanIDNumberData( _
    ByVal ws As Worksheet, _
    ByVal idCol As Integer, _
    ByVal statusCol As Integer)

    Dim lastRow As Long
    Dim r As Long
    Dim idStr As String
    Dim validCount As Long
    Dim invalidCount As Long

    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "資料不足，請確認工作表有標題列及資料列。", vbExclamation, "錯誤"
        Exit Sub
    End If

    ws.Cells(1, statusCol).Value = "驗證結果"
    ws.Cells(1, statusCol).Font.Bold = True

    validCount = 0
    invalidCount = 0

    For r = 2 To lastRow
        idStr = Trim(CStr(ws.Cells(r, idCol).Value))
        idStr = UCase(idStr)

        If IsValidTWID(idStr) Then
            ws.Cells(r, statusCol).Value = "有效"
            ws.Cells(r, statusCol).Font.Color = RGB(0, 128, 0)
            validCount = validCount + 1
        Else
            ws.Cells(r, statusCol).Value = "無效"
            ws.Cells(r, statusCol).Font.Color = RGB(200, 0, 0)
            ws.Cells(r, idCol).Interior.Color = RGB(255, 200, 200)
            invalidCount = invalidCount + 1
        End If
    Next r

    ws.Columns.AutoFit
    MsgBox "身分證字號驗證完成！" & vbCrLf & _
           "有效：" & validCount & " 筆" & vbCrLf & _
           "無效：" & invalidCount & " 筆", vbInformation, "清洗結果"
End Sub

' 範例使用入口
Sub TestCleanIDNumberData()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("身分證清洗")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "身分證清洗"
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "身分證字號"
    ws.Range("A1:B1").Font.Bold = True

    ws.Range("A2").Value = "王小明": ws.Range("B2").Value = "A123456789"
    ws.Range("A3").Value = "李美玲": ws.Range("B3").Value = "B234567890"
    ws.Range("A4").Value = "陳大華": ws.Range("B4").Value = "C12345678"
    ws.Range("A5").Value = "林志豪": ws.Range("B5").Value = "D987654321"
    ws.Range("A6").Value = "張雅慧": ws.Range("B6").Value = "1234567890"
    ws.Range("A7").Value = "黃俊賢": ws.Range("B7").Value = "F123456789"

    Call CleanIDNumberData(ws, 2, 3)
End Sub