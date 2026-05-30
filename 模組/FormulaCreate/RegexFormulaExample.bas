Attribute VB_Name = "RegexFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: RegexFormulaExample
'功能說明: 以VBA正規表達式輔助建立公式，進行資料驗證與萃取範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestRegexFormulaExample()
    Call CreateRegexFormulaSheet
End Sub

' 建立正規表達式公式範例工作表
Sub CreateRegexFormulaSheet()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreateRegexSheet(ThisWorkbook, "正規表達式公式範例")
    Call FillRegexSampleData(ws)
    Call ApplyRegexFormulas(ws)

    ws.Columns("A:E").AutoFit
    ws.Activate
    MsgBox "正規表達式公式範例已建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 填入範例來源資料
Private Sub FillRegexSampleData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "原始資料"
    ws.Range("B1").Value = "是否為Email"
    ws.Range("C1").Value = "是否為電話"
    ws.Range("D1").Value = "萃取數字"
    ws.Range("E1").Value = "清除非字母"
    ws.Range("A1:E1").Font.Bold = True
    ws.Range("A2").Value = "john@example.com"
    ws.Range("A3").Value = "0912-345-678"
    ws.Range("A4").Value = "ABC123XYZ"
    ws.Range("A5").Value = "user@test.org"
    ws.Range("A6").Value = "02-2345-6789"
    ws.Range("A7").Value = "Hello World 2024"
End Sub

' 套用正規表達式輔助公式（使用VBA函數填入結果）
Private Sub ApplyRegexFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    Dim sVal As String

    For i = 2 To 7
        sVal = CStr(ws.Cells(i, 1).Value)
        ws.Cells(i, 2).Value = CheckIsEmail(sVal)
        ws.Cells(i, 3).Value = CheckIsPhone(sVal)
        ws.Cells(i, 4).Value = ExtractNumbers(sVal)
        ws.Cells(i, 5).Value = RemoveNonAlphanumeric(sVal)
    Next i
End Sub

' 判斷是否為Email格式
Private Function CheckIsEmail(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"
    re.IgnoreCase = True
    If re.Test(s) Then
        CheckIsEmail = "是"
    Else
        CheckIsEmail = "否"
    End If
End Function

' 判斷是否為電話號碼格式
Private Function CheckIsPhone(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^0[0-9]{1,3}[\-]?[0-9]{3,4}[\-]?[0-9]{4}$"
    If re.Test(s) Then
        CheckIsPhone = "是"
    Else
        CheckIsPhone = "否"
    End If
End Function

' 萃取字串中的所有數字
Private Function ExtractNumbers(ByVal s As String) As String
    Dim re As Object
    Dim matches As Object
    Dim result As String
    Dim match As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[0-9]+"
    re.Global = True
    Set matches = re.Execute(s)
    result = ""
    For Each match In matches
        result = result & match.Value
    Next match
    ExtractNumbers = result
End Function

' 移除非字母數字字元
Private Function RemoveNonAlphanumeric(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[^a-zA-Z0-9]"
    re.Global = True
    RemoveNonAlphanumeric = re.Replace(s, "")
End Function

' 取得或建立工作表並清除內容
Private Function GetOrCreateRegexSheet(ByVal wb As Workbook, _
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
    Set GetOrCreateRegexSheet = ws
End Function
