Attribute VB_Name = "ClearPrintAreaFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearPrintAreaFormatting
'功能說明: 清除工作表的列印區域設定（PrintArea、PageBreaks）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點：清除目前工作表的列印區域相關設定
Sub TestClearPrintAreaFormatting()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Call ClearPrintAreaSettings(ws)
    MsgBox "列印區域設定已全部清除！", vbInformation, "完成"
End Sub

' 清除指定工作表的列印區域及分頁相關設定
' ws: 目標工作表
Sub ClearPrintAreaSettings(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    ws.PageSetup.PrintArea = ""

    Do While ws.HPageBreaks.Count > 0
        ws.HPageBreaks(1).Delete
    Loop

    Do While ws.VPageBreaks.Count > 0
        ws.VPageBreaks(1).Delete
    Loop

    With ws.PageSetup
        .Zoom = 100
        .FitToPagesWide = False
        .FitToPagesTall = False
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With

    Exit Sub

ErrorHandler:
    MsgBox "清除列印區域設定時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 清除活頁簿所有工作表的列印區域設定
Sub ClearAllSheetsPrintArea()
    On Error GoTo ErrorHandler

    Dim ws    As Worksheet
    Dim count As Long
    count = 0

    For Each ws In ThisWorkbook.Worksheets
        Call ClearPrintAreaSettings(ws)
        count = count + 1
    Next ws

    MsgBox "已清除 " & count & " 個工作表的列印區域設定。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "清除所有工作表列印設定時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 顯示目前工作表的列印區域資訊
Sub ShowCurrentPrintAreaInfo()
    Dim ws  As Worksheet
    Dim msg As String
    Set ws = ActiveSheet

    msg = "工作表：" & ws.Name & Chr(10)

    If ws.PageSetup.PrintArea = "" Then
        msg = msg & "列印區域：（未設定）" & Chr(10)
    Else
        msg = msg & "列印區域：" & ws.PageSetup.PrintArea & Chr(10)
    End If

    msg = msg & "水平分頁數：" & ws.HPageBreaks.Count & Chr(10)
    msg = msg & "垂直分頁數：" & ws.VPageBreaks.Count & Chr(10)
    msg = msg & "縮放比例：" & ws.PageSetup.Zoom & "%" & Chr(10)
    msg = msg & "列印標題列：" & ws.PageSetup.PrintTitleRows & Chr(10)
    msg = msg & "列印標題欄：" & ws.PageSetup.PrintTitleColumns

    MsgBox msg, vbInformation, "目前列印區域資訊"
End Sub
