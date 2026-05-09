Attribute VB_Name = "ClearFontFormattingOnly"
Option Explicit

'*************************************************************************************
'模組名稱: ClearFontFormattingOnly
'功能說明: 只清除字型格式（粗體、斜體、顏色、大小），保留框線/填色/數字格式
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 清除目前選取範圍的字型格式
Sub ClearFontFormattingOnly()
    Dim targetRange As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除字型格式的儲存格範圍。", vbExclamation, "警告"
        Exit Sub
    End If

    Set targetRange = Selection

    On Error GoTo ErrHandler

    With targetRange.Font
        .Name       = "新細明體"
        .Size       = 12
        .Bold       = False
        .Italic     = False
        .Underline  = xlUnderlineStyleNone
        .Strikethrough = False
        .Color      = RGB(0, 0, 0)
        .TintAndShade = 0
    End With

    MsgBox "已清除選取範圍的字型格式（共 " & targetRange.Cells.Count & " 個儲存格）。", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除字型格式時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 示範：建立含字型格式的資料再清除
Sub DemoClearFontFormatting()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    With ws.Range("A1:C3")
        .Value = "示範文字"
        .Font.Bold = True
        .Font.Italic = True
        .Font.Color = RGB(255, 0, 0)
        .Font.Size = 16
        .Font.Underline = xlUnderlineStyleSingle
    End With

    MsgBox "已建立帶有字型格式的資料，請選取 A1:C3 後執行 ClearFontFormattingOnly。", _
           vbInformation, "提示"
End Sub