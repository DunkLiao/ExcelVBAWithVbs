Option Explicit
'*************************************************************************************
'模組名稱: CleanBooleanData
'功能說明: 自動清理並標準化範圍中的布林/邏輯值（TRUE/FALSE、是/否、Y/N、1/0 等）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Function NormalizeBooleanValue(ByVal cellVal As Variant) As Variant
    ' 將常見布林表示方式轉換為標準 TRUE/FALSE
    ' 回傳 Variant：若可辨識則回傳 Boolean，否則回傳原值
    Dim strVal As String

    If IsEmpty(cellVal) Then
        NormalizeBooleanValue = cellVal
        Exit Function
    End If

    strVal = Trim(UCase(CStr(cellVal)))

    ' 對應 TRUE 的常見表示
    Select Case strVal
        Case "TRUE", "是", "Y", "YES", "1", "對", "V", "T", "OUI"
            NormalizeBooleanValue = True
            Exit Function
    End Select

    ' 對應 FALSE 的常見表示
    Select Case strVal
        Case "FALSE", "否", "N", "NO", "0", "錯", "X", "F", "NON"
            NormalizeBooleanValue = False
            Exit Function
    End Select

    ' 無法辨識，保留原值
    NormalizeBooleanValue = cellVal
End Function

Sub CleanBooleanData()
    ' 將選取範圍中的布林值標準化為 Excel TRUE/FALSE
    Dim rng As Range
    Dim cell As Range
    Dim originalVal As Variant
    Dim normalizedVal As Variant
    Dim cleanCount As Long
    Dim skipCount As Long

    On Error GoTo ErrHandler

    On Error Resume Next
    Set rng = Application.InputBox( _
        "請選取要清理布林值的範圍：", "選取範圍", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    cleanCount = 0
    skipCount = 0

    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            originalVal = cell.Value
            normalizedVal = NormalizeBooleanValue(originalVal)

            If VarType(normalizedVal) = vbBoolean Then
                ' 值有變化時才更新
                If VarType(originalVal) <> vbBoolean Or (CBool(originalVal) <> CBool(normalizedVal)) Then
                    cell.Value = normalizedVal
                    cell.Interior.Color = RGB(198, 239, 206)  ' 淺綠標示已清理
                    cleanCount = cleanCount + 1
                End If
            Else
                ' 無法辨識的值
                skipCount = skipCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "布林值清理完成！" & vbNewLine & _
           "已標準化：" & cleanCount & " 個儲存格" & vbNewLine & _
           "無法辨識（保留原值）：" & skipCount & " 個儲存格", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

Sub CleanBooleanDataToChineseText()
    ' 將布林值標準化後轉換為繁體中文「是/否」文字格式
    Dim rng As Range
    Dim cell As Range
    Dim normalizedVal As Variant
    Dim cleanCount As Long

    On Error GoTo ErrHandler

    On Error Resume Next
    Set rng = Application.InputBox( _
        "請選取要轉換為「是/否」的範圍：", "選取範圍", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    cleanCount = 0

    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            normalizedVal = NormalizeBooleanValue(cell.Value)
            If VarType(normalizedVal) = vbBoolean Then
                cell.Value = IIf(CBool(normalizedVal), "是", "否")
                cell.Interior.Color = RGB(221, 235, 247)  ' 淺藍標示
                cleanCount = cleanCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "布林值轉換完成！共轉換 " & cleanCount & " 個儲存格為「是/否」格式。", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
