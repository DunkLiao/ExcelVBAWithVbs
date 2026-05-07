Attribute VB_Name = "ClearAllSheetsFormatting"
Option Explicit

' ============================================================
' 範例：一次清除活頁簿所有工作表的儲存格格式
' 功能：保留資料內容，僅清除字型、色彩、框線等格式設定
' ============================================================
Sub ClearAllSheetsFormatting()
    Dim ws      As Worksheet
    Dim intCnt  As Integer

    On Error GoTo ErrHandler

    Dim blnConfirm As Boolean
    blnConfirm = (MsgBox( _
        "確定要清除活頁簿所有工作表的儲存格格式？" & Chr(10) & _
        "（此操作將清除字型、色彩、框線等，但保留資料內容）", _
        vbYesNo + vbQuestion, "確認清除格式") = vbYes)

    If Not blnConfirm Then
        MsgBox "操作已取消。", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    intCnt = 0

    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.ClearFormats
        intCnt = intCnt + 1
    Next ws

    Application.ScreenUpdating = True
    MsgBox "已清除 " & intCnt & " 個工作表的所有儲存格格式。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
