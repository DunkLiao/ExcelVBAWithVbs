Attribute VB_Name = "ClearValidationFormatting"
Option Explicit

' ============================================================
' 模組名稱：ClearValidationFormatting
' 功能說明：清除工作表中指定範圍或整張工作表的資料驗證設定
'           提供：選取範圍清除 / 整張工作表清除 / 僅清除含下拉選單的驗證
' ============================================================

Sub ClearValidationFormatting()
    Dim choice As String
    
    choice = InputBox("請選擇清除模式：" & vbCrLf & _
                      "1 = 清除選取範圍的資料驗證" & vbCrLf & _
                      "2 = 清除整張工作表的資料驗證" & vbCrLf & _
                      "3 = 僅清除含下拉選單的驗證", _
                      "清除資料驗證", "1")
    
    If choice = "" Then
        MsgBox "已取消操作。", vbInformation, "取消"
        Exit Sub
    End If
    
    Select Case choice
        Case "1"
            Call ClearValidationInRange
        Case "2"
            Call ClearValidationEntireSheet
        Case "3"
            Call ClearDropdownValidationOnly
        Case Else
            MsgBox "請輸入 1、2 或 3。", vbExclamation, "輸入錯誤"
    End Select
End Sub

' 清除選取範圍的資料驗證
Private Sub ClearValidationInRange()
    Dim rng As Range
    Dim count As Long
    
    On Error GoTo ErrHandler
    
    Set rng = Application.InputBox( _
        "請選取要清除資料驗證的範圍：", "選取範圍", Type:=8)
    
    If rng Is Nothing Then Exit Sub
    
    count = 0
    
    Dim cell As Range
    For Each cell In rng
        On Error Resume Next
        If cell.Validation.Type >= 0 Then
            count = count + 1
        End If
        On Error GoTo ErrHandler
    Next cell
    
    rng.Validation.Delete
    
    MsgBox "已清除選取範圍的資料驗證。" & vbCrLf & _
           "受影響範圍：" & rng.Address, vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

' 清除整張工作表的資料驗證
Private Sub ClearValidationEntireSheet()
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = ActiveSheet
    
    Dim ans As VbMsgBoxResult
    ans = MsgBox("確定要清除「" & ws.Name & "」整張工作表的所有資料驗證嗎？", _
                 vbQuestion + vbYesNo, "確認")
    If ans = vbNo Then Exit Sub
    
    ws.Cells.Validation.Delete
    
    MsgBox "已清除整張工作表「" & ws.Name & "」的所有資料驗證。", _
           vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

' 僅清除含下拉選單（清單型）的驗證
Private Sub ClearDropdownValidationOnly()
    Dim ws          As Worksheet
    Dim cell        As Range
    Dim clearCount  As Long
    Dim rngAll      As Range
    
    On Error GoTo ErrHandler
    
    Set ws = ActiveSheet
    clearCount = 0
    
    Application.ScreenUpdating = False
    
    ' 取得所有有驗證的儲存格
    On Error Resume Next
    Set rngAll = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    On Error GoTo ErrHandler
    
    If rngAll Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "此工作表沒有任何資料驗證。", vbInformation, "提示"
        Exit Sub
    End If
    
    ' 逐一檢查是否為下拉清單（xlValidateList = 3）
    For Each cell In rngAll
        On Error Resume Next
        If cell.Validation.Type = xlValidateList Then
            cell.Validation.Delete
            clearCount = clearCount + 1
        End If
        On Error GoTo ErrHandler
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "已清除 " & clearCount & " 個含下拉選單的資料驗證儲存格。", _
           vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub