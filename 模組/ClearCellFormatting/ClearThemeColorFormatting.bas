Option Explicit
Attribute VB_Name = "ClearThemeColorFormatting"
'*************************************************************************************
'模組名稱: ClearThemeColorFormatting
'功能說明: 清除儲存格套用的佈景主題色彩格式，還原為無填滿與標準字型顏色
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestClearThemeColorFormatting()
    Call ClearThemeColors(ActiveSheet)
End Sub

Sub ClearThemeColors(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    Call ApplyThemeColorsDemo(ws)
    MsgBox "已套用示範佈景主題色彩，即將清除...", vbInformation, "示範"

    Call RemoveThemeColorFormatting(ws.UsedRange)

    MsgBox "佈景主題色彩格式已清除完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "清除佈景主題色彩時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub ApplyThemeColorsDemo(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:E1").Interior.ThemeColor = xlThemeColorAccent1
    ws.Range("A1:E1").Font.ThemeColor = xlThemeColorLight1
    ws.Range("A2:E2").Interior.ThemeColor = xlThemeColorAccent2
    ws.Range("A3:E3").Interior.ThemeColor = xlThemeColorAccent3
    ws.Range("A4:E4").Interior.ThemeColor = xlThemeColorAccent4

    Dim r As Integer
    For r = 1 To 4
        ws.Cells(r, 1).Value = "列 " & r & " 佈景主題色彩示範"
    Next r
End Sub

Private Sub RemoveThemeColorFormatting(ByVal targetRange As Range)
    On Error Resume Next
    Dim c As Range
    For Each c In targetRange.Cells
        c.Interior.ColorIndex = xlNone
        c.Font.ColorIndex = xlAutomatic
    Next c
    On Error GoTo 0
End Sub
