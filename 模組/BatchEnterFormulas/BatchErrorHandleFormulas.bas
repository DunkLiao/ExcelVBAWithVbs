Attribute VB_Name = "BatchErrorHandleFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchErrorHandleFormulas
'功能說明: 批次將選取範圍內的公式包裝為IFERROR或IFNA，避免顯示錯誤值
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub BatchWrapWithIFERROR()
    Dim targetRange As Range
    Dim cell As Range
    Dim originalFormula As String
    Dim errorValue As String
    Dim processCount As Integer

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取含公式的儲存格範圍。", vbExclamation, "提示"
        Exit Sub
    End If

    errorValue = InputBox("請輸入錯誤時顯示的替代值（預設空白，留空即可）：", "設定替代值", "")

    Set targetRange = Selection
    processCount = 0

    Application.ScreenUpdating = False

    For Each cell In targetRange
        If cell.HasFormula Then
            originalFormula = Mid(cell.Formula, 2)
            If Left(UCase(originalFormula), 8) <> "IFERROR(" Then
                Dim q As String
                q = Chr(34)
                If errorValue = "" Then
                    cell.Formula = "=IFERROR(" & originalFormula & "," & q & q & ")"
                Else
                    cell.Formula = "=IFERROR(" & originalFormula & "," & errorValue & ")"
                End If
                processCount = processCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "已包裝 " & processCount & " 個公式為 IFERROR 格式。", vbInformation, "完成"
End Sub

Sub BatchWrapWithIFNA()
    Dim cell As Range
    Dim processCount As Integer

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取含公式的儲存格範圍。", vbExclamation, "提示"
        Exit Sub
    End If

    processCount = 0
    Application.ScreenUpdating = False

    For Each cell In Selection
        If cell.HasFormula Then
            Dim f As String
            f = Mid(cell.Formula, 2)
            If Left(UCase(f), 5) <> "IFNA(" Then
                cell.Formula = "=IFNA(" & f & ",""查無資料"")"
                processCount = processCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "已包裝 " & processCount & " 個公式為 IFNA 格式。", vbInformation, "完成"
End Sub
