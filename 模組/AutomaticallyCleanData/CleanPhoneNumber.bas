Attribute VB_Name = "CleanPhoneNumber"
Option Explicit

' ============================================================
' 模組名稱：CleanPhoneNumber
' 功能說明：清理並標準化電話號碼格式
'           支援台灣手機號碼（09xxxxxxxx）
'           與市話（區碼-號碼）的格式統一
' 使用方式：選取包含電話號碼的欄位後執行
' ============================================================

Sub CleanPhoneNumber()
    Dim rng         As Range
    Dim cell        As Range
    Dim original    As String
    Dim cleaned     As String
    Dim fixedCount  As Long
    Dim skipCount   As Long
    
    On Error GoTo ErrHandler
    
    ' 選取要清理的範圍
    Set rng = Application.InputBox( _
        "請選取包含電話號碼的儲存格範圍：", "選取範圍", Type:=8)
    
    If rng Is Nothing Then
        MsgBox "已取消操作。", vbInformation, "取消"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    fixedCount = 0
    skipCount = 0
    
    For Each cell In rng
        original = Trim(CStr(cell.Value))
        
        If original = "" Then
            skipCount = skipCount + 1
        Else
            cleaned = NormalizePhoneNumber(original)
            
            If cleaned <> original Then
                cell.Value = cleaned
                ' 標示已修改的儲存格（淺藍底）
                cell.Interior.Color = RGB(173, 216, 230)
                fixedCount = fixedCount + 1
            ElseIf cleaned = "格式錯誤" Then
                ' 標示無法識別的格式（橘色底）
                cell.Interior.Color = RGB(255, 165, 0)
                skipCount = skipCount + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "電話號碼清理完成！" & vbCrLf & _
           "已修正：" & fixedCount & " 筆" & vbCrLf & _
           "無法識別（跳過）：" & skipCount & " 筆" & vbCrLf & _
           "（淺藍色 = 已修正，橘色 = 格式無法識別）", _
           vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

' 電話號碼標準化核心函式
Private Function NormalizePhoneNumber(ByVal phone As String) As String
    Dim result  As String
    Dim digits  As String
    Dim i       As Integer
    Dim c       As String
    
    ' 移除常見分隔字元（空格、括號、連字號、加號）
    result = phone
    result = Replace(result, " ", "")
    result = Replace(result, "-", "")
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    result = Replace(result, "+", "")
    result = Replace(result, ".", "")
    result = Replace(result, "#", "")
    
    ' 移除分機號碼（以 # 或 ext 開頭的部分已先去除 #，這裡處理 ext）
    Dim extPos As Integer
    extPos = InStr(LCase(result), "ext")
    If extPos > 0 Then result = Left(result, extPos - 1)
    
    ' 提取純數字
    digits = ""
    For i = 1 To Len(result)
        c = Mid(result, i, 1)
        If c >= "0" And c <= "9" Then
            digits = digits & c
        End If
    Next i
    
    ' 判斷並格式化
    Select Case Len(digits)
        Case 10
            If Left(digits, 2) = "09" Then
                ' 台灣手機：09xx-xxxxxx
                NormalizePhoneNumber = Left(digits, 4) & "-" & Mid(digits, 5)
            ElseIf Left(digits, 1) = "0" Then
                ' 市話（含區碼，共10碼）：02-xxxxxxxx / 03~09-xxxxxxx
                Dim areaCode As String
                If Left(digits, 2) = "02" Then
                    ' 台北（02 + 8碼）
                    areaCode = "02"
                    NormalizePhoneNumber = areaCode & "-" & Mid(digits, 3)
                ElseIf Left(digits, 3) >= "037" Then
                    ' 其他區碼（03x + 7碼）
                    areaCode = Left(digits, 3)
                    NormalizePhoneNumber = areaCode & "-" & Mid(digits, 4)
                Else
                    NormalizePhoneNumber = Left(digits, 2) & "-" & Mid(digits, 3)
                End If
            Else
                NormalizePhoneNumber = digits
            End If
        Case 9
            If Left(digits, 2) = "09" Then
                ' 去掉前導 0 的手機（9碼）
                NormalizePhoneNumber = "0" & Left(digits, 3) & "-" & Mid(digits, 4)
            Else
                ' 市話無區碼（可能是7碼或8碼）
                NormalizePhoneNumber = digits
            End If
        Case 8
            ' 市話號碼（無區碼）：xxxx-xxxx
            NormalizePhoneNumber = Left(digits, 4) & "-" & Mid(digits, 5)
        Case 7
            ' 市話號碼（無區碼，7碼）：xxx-xxxx
            NormalizePhoneNumber = Left(digits, 3) & "-" & Mid(digits, 4)
        Case 0
            NormalizePhoneNumber = "格式錯誤"
        Case Else
            NormalizePhoneNumber = digits
    End Select
End Function