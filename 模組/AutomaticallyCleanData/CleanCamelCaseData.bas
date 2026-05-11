Attribute VB_Name = "CleanCamelCaseData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanCamelCaseData
'功能說明: 將選取範圍內的駝峰式（camelCase/PascalCase）文字轉換為空格分隔格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Private Function SplitCamelCase(ByVal text As String) As String
    Dim result As String
    Dim i As Integer
    Dim c As String
    Dim prevChar As String
    Dim isUpper As Boolean
    Dim prevIsLower As Boolean

    result = ""
    prevChar = ""

    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        If i > 1 Then
            isUpper = (c >= "A" And c <= "Z")
            prevIsLower = (prevChar >= "a" And prevChar <= "z")
            If isUpper And prevIsLower Then
                result = result & " "
            End If
        End If
        result = result & c
        prevChar = c
    Next i

    SplitCamelCase = result
End Function

Sub ConvertCamelCaseToSpaced()
    Dim cell As Range
    Dim processCount As Integer
    Dim original As String
    Dim converted As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取含駝峰式文字的儲存格範圍。", vbExclamation, "提示"
        Exit Sub
    End If

    processCount = 0
    Application.ScreenUpdating = False

    For Each cell In Selection
        If Not IsEmpty(cell) And Not cell.HasFormula Then
            original = CStr(cell.Value)
            converted = SplitCamelCase(original)
            If converted <> original Then
                cell.Value = converted
                processCount = processCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "已轉換 " & processCount & " 個駝峰式字串為空格分隔格式。", vbInformation, "完成"
End Sub

Sub ConvertCamelCaseToPascalCase()
    Dim cell As Range
    Dim processCount As Integer
    Dim val As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取儲存格範圍。", vbExclamation, "提示"
        Exit Sub
    End If

    processCount = 0
    Application.ScreenUpdating = False

    For Each cell In Selection
        If Not IsEmpty(cell) And Not cell.HasFormula Then
            val = CStr(cell.Value)
            If Len(val) > 0 Then
                cell.Value = UCase(Left(val, 1)) & Mid(val, 2)
                processCount = processCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "已轉換 " & processCount & " 個字串為 PascalCase。", vbInformation, "完成"
End Sub
