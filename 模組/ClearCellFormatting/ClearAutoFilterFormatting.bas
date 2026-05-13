Attribute VB_Name = "ClearAutoFilterFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: 清除自動篩選格式
'功能說明: 清除工作表上所有自動篩選設定，並還原所有隱藏的列
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub ClearAutoFilterFormatting()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Call RemoveAutoFilter(ws)
End Sub

Sub ClearAllSheetsAutoFilter()
    Dim ws As Worksheet
    Dim cnt As Integer
    cnt = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode Then
            Call RemoveAutoFilter(ws)
            cnt = cnt + 1
        End If
    Next ws

    MsgBox "已清除 " & cnt & " 張工作表的自動篩選設定。", vbInformation, "完成"
End Sub

Private Sub RemoveAutoFilter(ByVal ws As Worksheet)
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
        ws.AutoFilterMode = False
    End If
    ws.Rows.Hidden = False
    MsgBox "工作表「" & ws.Name & "」的自動篩選已清除，所有列已顯示。", _
           vbInformation, "完成"
End Sub

Sub ClearFilterCriteriaOnly()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
            MsgBox "篩選條件已清除，自動篩選下拉選單仍保留。", vbInformation, "完成"
        Else
            MsgBox "目前沒有套用任何篩選條件。", vbInformation, "提示"
        End If
    Else
        MsgBox "工作表目前未啟用自動篩選。", vbInformation, "提示"
    End If
End Sub
