Option Explicit

' 只清除選取範圍的條件式格式，保留字型、框線、底色與數值。
Public Sub ClearConditionalFormatsOnlyExample()
    On Error GoTo ErrHandler

    Dim targetRange As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除條件式格式的範圍。", vbExclamation, "提醒"
        Exit Sub
    End If

    Set targetRange = Selection
    targetRange.FormatConditions.Delete

    MsgBox "已清除選取範圍的條件式格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除條件式格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub