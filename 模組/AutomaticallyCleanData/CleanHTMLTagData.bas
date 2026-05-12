Option Explicit
Attribute VB_Name = "CleanHTMLTagData"
'*************************************************************************************
'模組名稱: CleanHTMLTagData
'功能說明: 自動清除儲存格內容中的 HTML 標籤，只保留純文字內容
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Function RemoveHTMLTags(ByVal inputText As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .Pattern = "<[^>]+(>|$)"
    End With
    RemoveHTMLTags = regEx.Replace(inputText, "")
    Set regEx = Nothing
End Function

Sub CleanHTMLTagData()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim cleanedCount As Long
    Dim original As String
    Dim cleaned As String

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.ActiveSheet
    cleanedCount = 0
    Set rng = ws.UsedRange

    For Each cell In rng
        If cell.Value <> "" Then
            original = CStr(cell.Value)
            If InStr(original, "<") > 0 And InStr(original, ">") > 0 Then
                cleaned = RemoveHTMLTags(original)
                If cleaned <> original Then
                    cell.Value = cleaned
                    cleanedCount = cleanedCount + 1
                End If
            End If
        End If
    Next cell

    MsgBox "HTML 標籤清除完成！共清理 " & cleanedCount & " 個儲存格。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "清除 HTML 標籤失敗"
End Sub