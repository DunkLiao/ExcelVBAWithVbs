'*************************************************************************************
'模組名稱: ClearAllFormattingInWorkbook
'功能說明: 清除活頁簿中所有工作表的儲存格格式設定（保留資料）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************
Option Explicit

Sub ClearAllFormattingInWorkbook()
    Dim ws          As Worksheet
    Dim usedRng     As Range
    Dim count       As Integer
    Dim confirm     As VbMsgBoxResult

    confirm = MsgBox("此操作將清除所有工作表的格式設定（資料不受影響）。" & vbCrLf & _
                     "確定要繼續嗎？", vbYesNo + vbExclamation, "確認")
    If confirm = vbNo Then Exit Sub

    count = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Visible = xlSheetVisible Then
            Set usedRng = ws.UsedRange
            If Not usedRng Is Nothing Then
                usedRng.ClearFormats
                count = count + 1
            End If
        End If
    Next ws

    MsgBox "已清除 " & count & " 個工作表的所有格式設定。", vbInformation, "完成"
End Sub
