Attribute VB_Name = "CleanTextCaseData"
Option Explicit

'*************************************************************************************
'模組名稱: CleanTextCaseData
'功能說明: 批次標準化選取範圍內的英文大小寫
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub CleanTextCaseData()
    Dim rng As Range
    Dim cell As Range
    Dim caseMode As Integer
    Dim msg As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要轉換大小寫的儲存格範圍！", vbExclamation
        Exit Sub
    End If

    Set rng = Selection

    msg = "請選擇大小寫轉換模式：" & vbNewLine & _
        "1 = 全部大寫（UPPER CASE）" & vbNewLine & _
        "2 = 全部小寫（lower case）" & vbNewLine & _
        "3 = 首字大寫（Proper Case）"

    caseMode = CInt(InputBox(msg, "大小寫轉換", "1"))

    If caseMode < 1 Or caseMode > 3 Then
        MsgBox "無效的模式，請輸入 1、2 或 3！", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For Each cell In rng
        If Not IsEmpty(cell) And Not IsNumeric(cell.Value) Then
            Select Case caseMode
                Case 1
                    cell.Value = UCase(cell.Value)
                Case 2
                    cell.Value = LCase(cell.Value)
                Case 3
                    cell.Value = StrConv(CStr(cell.Value), vbProperCase)
            End Select
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "大小寫轉換完成！", vbInformation
End Sub
