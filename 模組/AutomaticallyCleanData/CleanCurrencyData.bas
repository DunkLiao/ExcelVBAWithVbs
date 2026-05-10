'*************************************************************************************
'模組名稱: CleanCurrencyData
'功能說明: 自動清理儲存格中的貨幣符號、千分位符號，轉換為純數值
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************
Option Explicit

Sub CleanCurrencyData()
    Dim ws          As Worksheet
    Dim rng         As Range
    Dim cell        As Range
    Dim cleaned     As String
    Dim cleanCount  As Long

    Set ws = ActiveSheet

    ' 讓使用者選擇要清理的範圍
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="請選擇要清理貨幣格式的範圍：", _
        Title:="選擇範圍", _
        Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    cleanCount = 0
    Application.ScreenUpdating = False

    For Each cell In rng
        If cell.Value <> "" And Not IsEmpty(cell.Value) Then
            cleaned = CStr(cell.Value)
            ' 移除常見貨幣符號
            cleaned = Replace(cleaned, "$", "")
            cleaned = Replace(cleaned, "NT$", "")
            cleaned = Replace(cleaned, Chr(165), "")  ' ￥
            cleaned = Replace(cleaned, Chr(8364), "") ' EUR sign fallback
            ' 移除千分位符號
            cleaned = Replace(cleaned, ",", "")
            ' 移除空白
            cleaned = Trim(cleaned)
            ' 若可轉換為數值則更新
            If IsNumeric(cleaned) Then
                cell.Value = CDbl(cleaned)
                cell.NumberFormat = "#,##0.00"
                cleanCount = cleanCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "貨幣資料清理完成，共處理 " & cleanCount & " 個儲存格。", vbInformation, "完成"
End Sub
