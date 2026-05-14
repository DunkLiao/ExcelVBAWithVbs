Attribute VB_Name = "SplitByRegexPattern"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByRegexPattern
'功能說明: 依正規表示式（RegExp）比對結果，將工作表資料列分割至不同工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口：依第一欄是否符合 Email 格式分割
Sub TestSplitByRegex()
    Dim ws As Worksheet
    Set ws = GetOrCreateRegexWs(ThisWorkbook, "正規分割測試")
    ws.Cells.Clear

    ' 填入測試資料
    ws.Range("A1").Value = "名稱"
    ws.Range("B1").Value = "聯絡資訊"
    ws.Range("A2").Value = "王小明"
    ws.Range("B2").Value = "wang@example.com"
    ws.Range("A3").Value = "李大華"
    ws.Range("B3").Value = "0912345678"
    ws.Range("A4").Value = "陳美玲"
    ws.Range("B4").Value = "chen@mail.tw"
    ws.Range("A5").Value = "林俊傑"
    ws.Range("B5").Value = "02-12345678"
    ws.Range("A6").Value = "張志遠"
    ws.Range("B6").Value = "zhang@corp.com"

    Call SplitByRegexPattern(ws, 2, "[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", _
        "Email符合", "Email不符")
End Sub

' 依正規表示式分割工作表資料
' ws        : 來源工作表
' colIndex  : 要比對的欄位索引（1 = A 欄）
' pattern   : 正規表示式字串
' matchName : 符合條件的目標工作表名稱
' noMatchName: 不符合條件的目標工作表名稱
Sub SplitByRegexPattern(ByVal ws As Worksheet, ByVal colIndex As Long, _
    ByVal pattern As String, ByVal matchName As String, ByVal noMatchName As String)
    On Error GoTo ErrorHandler

    Dim re As Object
    Dim wsMatch As Worksheet
    Dim wsNoMatch As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim matchRow As Long
    Dim noMatchRow As Long
    Dim cellValue As String

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = False

    Set wsMatch = GetOrCreateRegexWs(ThisWorkbook, matchName)
    Set wsNoMatch = GetOrCreateRegexWs(ThisWorkbook, noMatchName)
    wsMatch.Cells.Clear
    wsNoMatch.Cells.Clear

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' 複製標題列
    ws.Rows(1).Resize(1, lastCol).Copy wsMatch.Range("A1")
    ws.Rows(1).Resize(1, lastCol).Copy wsNoMatch.Range("A1")
    matchRow = 2
    noMatchRow = 2

    For i = 2 To lastRow
        cellValue = CStr(ws.Cells(i, colIndex).Value)
        If re.Test(cellValue) Then
            ws.Rows(i).Resize(1, lastCol).Copy wsMatch.Cells(matchRow, 1)
            matchRow = matchRow + 1
        Else
            ws.Rows(i).Resize(1, lastCol).Copy wsNoMatch.Cells(noMatchRow, 1)
            noMatchRow = noMatchRow + 1
        End If
    Next i

    wsMatch.UsedRange.Columns.AutoFit
    wsNoMatch.UsedRange.Columns.AutoFit

    MsgBox "分割完成！" & vbCrLf & _
        "符合：" & matchName & "（" & matchRow - 2 & " 列）" & vbCrLf & _
        "不符合：" & noMatchName & "（" & noMatchRow - 2 & " 列）", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "分割時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateRegexWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateRegexWs = ws
End Function
