Option Explicit
Attribute VB_Name = "CleanDuplicateSpaces"
'*************************************************************************************
'模組名稱: 清除重複空格
'功能說明: 掃描選取範圍的文字儲存格，將多個連續空格壓縮為單一空格，
'          並清除前後多餘空格
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub CleanDuplicateSpaces()
    On Error GoTo ErrorHandler

    Dim sel As Range
    On Error Resume Next
    Set sel = Selection
    On Error GoTo ErrorHandler

    If sel Is Nothing Then
        MsgBox "請先選取要清理的範圍。", vbExclamation, "提示"
        Exit Sub
    End If

    Dim cell As Range
    Dim cleanedCount As Long
    Dim original As String
    Dim cleaned As String
    cleanedCount = 0

    Application.ScreenUpdating = False

    For Each cell In sel.Cells
        If cell.HasFormula Then GoTo NextCell
        If Not IsEmpty(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                original = cell.Value
                cleaned = RemoveDuplicateSpaces(original)
                If cleaned <> original Then
                    cell.Value = cleaned
                    cleanedCount = cleanedCount + 1
                End If
            End If
        End If
NextCell:
    Next cell

    Application.ScreenUpdating = True
    MsgBox "清理完成！共修改了 " & cleanedCount & " 個儲存格。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清除重複空格時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function RemoveDuplicateSpaces(ByVal s As String) As String
    ' 先去除前後空格
    s = Trim(s)
    ' 壓縮連續空格為單一空格
    Do While InStr(s, "  ") > 0
        s = Join(Split(s, "  "), " ")
    Loop
    RemoveDuplicateSpaces = s
End Function
