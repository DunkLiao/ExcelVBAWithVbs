Attribute VB_Name = "CleanBarcodeScanData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanBarcodeScanData
'功能說明: 自動清理條碼掃描機輸入的資料，去除換行符號、前後空白及非標準字元
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestCleanBarcodeScanData()
    Call CleanBarcodeData
End Sub

' 清理條碼掃描資料
Sub CleanBarcodeData()
    Dim ws As Worksheet
    Dim lngLastRow As Long
    Dim i As Long
    Dim sRaw As String
    Dim sClean As String
    Dim intFixed As Integer

    On Error GoTo ErrHandler
    Set ws = GetOrCreateBarcodeSheet(ThisWorkbook, "條碼資料清理")
    Call FillBarcodeSampleData(ws)

    lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    intFixed = 0

    Application.ScreenUpdating = False

    ws.Range("C1").Value = "清理後條碼"
    ws.Range("D1").Value = "是否修改"
    ws.Range("E1").Value = "條碼長度"
    ws.Range("C1:E1").Font.Bold = True

    For i = 2 To lngLastRow
        sRaw = CStr(ws.Cells(i, 2).Value)
        sClean = CleanBarcodeString(sRaw)
        ws.Cells(i, 3).Value = sClean
        ws.Cells(i, 5).Value = Len(sClean)
        If sRaw <> sClean Then
            ws.Cells(i, 4).Value = "已修改"
            ws.Cells(i, 4).Font.Color = RGB(255, 0, 0)
            ws.Cells(i, 3).Font.Color = RGB(0, 128, 0)
            intFixed = intFixed + 1
        Else
            ws.Cells(i, 4).Value = "無需修改"
            ws.Cells(i, 4).Font.Color = RGB(128, 128, 128)
        End If
    Next i

    ws.Columns("A:E").AutoFit
    ws.Activate
    Application.ScreenUpdating = True
    MsgBox "條碼資料清理完成！共修正 " & intFixed & " 筆資料。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清理單一條碼字串
Private Function CleanBarcodeString(ByVal s As String) As String
    Dim result As String
    Dim cleaned As String
    Dim i As Integer
    Dim c As String
    Dim code As Integer

    result = s
    result = Replace(result, Chr(13), "")
    result = Replace(result, Chr(10), "")
    result = Replace(result, Chr(9), "")
    result = Trim(result)

    cleaned = ""
    For i = 1 To Len(result)
        c = Mid(result, i, 1)
        code = Asc(c)
        If (code >= 48 And code <= 57) Or _
           (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or _
           code = 45 Or code = 95 Then
            cleaned = cleaned & c
        End If
    Next i

    CleanBarcodeString = UCase(cleaned)
End Function

' 填入條碼掃描範例資料
Private Sub FillBarcodeSampleData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "序號"
    ws.Range("B1").Value = "原始條碼"
    ws.Range("A1:B1").Font.Bold = True

    ws.Range("A2").Value = 1
    ws.Range("B2").Value = "ABC-12345" & Chr(13) & Chr(10)

    ws.Range("A3").Value = 2
    ws.Range("B3").Value = "  xyz-67890  "

    ws.Range("A4").Value = 3
    ws.Range("B4").Value = "DEF-11111"

    ws.Range("A5").Value = 4
    ws.Range("B5").Value = "ghi" & Chr(9) & "22222"

    ws.Range("A6").Value = 5
    ws.Range("B6").Value = "JKL-33333" & Chr(0)

    ws.Range("A7").Value = 6
    ws.Range("B7").Value = "mno-44444"

    ws.Range("A8").Value = 7
    ws.Range("B8").Value = "PQR-55555"

    ws.Columns("A:B").AutoFit
End Sub

' 取得或建立工作表並清除內容
Private Function GetOrCreateBarcodeSheet(ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateBarcodeSheet = ws
End Function
