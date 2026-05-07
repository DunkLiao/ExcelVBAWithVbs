Attribute VB_Name = "StandardizeTextData"
Option Explicit

' ============================================================
' 範例：自動標準化文字資料（去空白、統一大小寫、全半形轉換）
' 功能：對選取範圍套用 Trim、大小寫標準化，並移除全形空白
' ============================================================
Sub StandardizeSelectedTextData()
    Dim rng     As Range
    Dim cell    As Range
    Dim strVal  As String
    Dim intCnt  As Integer

    On Error GoTo ErrHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要標準化的儲存格範圍。", vbExclamation
        Exit Sub
    End If

    Set rng = Selection
    Application.ScreenUpdating = False
    intCnt = 0

    For Each cell In rng
        If Not IsEmpty(cell) And Not IsNumeric(cell.Value) Then
            strVal = cell.Value
            ' 去除前後空白
            strVal = Trim(strVal)
            ' 移除全形空白（Unicode 12288）
            strVal = Replace(strVal, Chr(12288), "")
            ' 移除多餘半形空白
            Do While InStr(strVal, "  ") > 0
                strVal = Replace(strVal, "  ", " ")
            Loop
            If cell.Value <> strVal Then
                cell.Value = strVal
                intCnt = intCnt + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "文字標準化完成，共修正 " & intCnt & " 個儲存格。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
