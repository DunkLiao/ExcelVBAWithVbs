Option Explicit
Attribute VB_Name = "ClearAllCellStyles"
'*************************************************************************************
'模組名稱: ClearAllCellStyles
'功能說明: 清除作用中活頁簿中所有儲存格的自訂樣式，還原為預設格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestClearAllCellStyles()
    Call ClearAllCellStylesInWorkbook
End Sub

' 清除活頁簿中所有自訂儲存格樣式
Sub ClearAllCellStylesInWorkbook()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim style As Style
    Dim styleName As String
    Dim count As Long
    Dim msg As String

    Set wb = ThisWorkbook
    count = 0

    ' 逐一檢查並刪除非內建樣式
    For Each style In wb.Styles
        styleName = style.Name
        If Not style.BuiltIn Then
            On Error Resume Next
            style.Delete
            If Err.Number = 0 Then
                count = count + 1
            End If
            On Error GoTo ErrorHandler
        End If
    Next style

    ' 對每個工作表的已用範圍清除所有格式
    Application.ScreenUpdating = False
    For Each ws In wb.Worksheets
        If ws.UsedRange.Cells.Count > 0 Then
            ws.UsedRange.ClearFormats
        End If
    Next ws
    Application.ScreenUpdating = True

    msg = "已清除 " & count & " 個自訂樣式" & vbCrLf
    msg = msg & "並清除所有工作表的儲存格格式。"
    MsgBox msg, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
