Option Explicit
Attribute VB_Name = "ClearPageBreakFormatting"
'*************************************************************************************
'模組名稱: ClearPageBreakFormatting
'功能說明: 清除工作表中所有手動設定的水平與垂直分頁符號
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub ClearPageBreakFormatting()
    Dim ws As Worksheet
    Dim hpbCount As Integer
    Dim vpbCount As Integer
    Dim vpb As VPageBreak
    Dim msg As String

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.ActiveSheet

    hpbCount = ws.HPageBreaks.Count
    vpbCount = 0

    For Each vpb In ws.VPageBreaks
        If vpb.Type = xlPageBreakManual Then
            vpbCount = vpbCount + 1
        End If
    Next vpb

    ws.ResetAllPageBreaks

    msg = "已清除「" & ws.Name & "」的手動分頁符號：" & vbCrLf
    msg = msg & "  水平分頁符號：" & hpbCount & " 個" & vbCrLf
    msg = msg & "  垂直分頁符號：" & vpbCount & " 個"

    MsgBox msg, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "清除分頁符號失敗"
End Sub