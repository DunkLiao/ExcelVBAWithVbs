Option Explicit
Attribute VB_Name = "HighlightKeywordFormatting"
'*************************************************************************************
'模組名稱: 關鍵字醒目標示格式
'功能說明: 對工作表指定範圍套用條件式格式，含有指定關鍵字的儲存格
'          自動填滿黃色背景並加粗文字
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub TestHighlightKeywordFormatting()
    Call ApplyHighlightKeywordFormatting(ActiveSheet, ActiveSheet.UsedRange, "緊急")
End Sub

Sub ApplyHighlightKeywordFormatting( _
    ByVal ws As Worksheet, _
    ByVal targetRange As Range, _
    ByVal keyword As String)

    On Error GoTo ErrorHandler

    If targetRange Is Nothing Then
        MsgBox "請指定有效範圍。", vbExclamation, "錯誤"
        Exit Sub
    End If

    If keyword = "" Then
        MsgBox "請指定關鍵字。", vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 清除原有條件格式
    targetRange.FormatConditions.Delete

    ' 新增包含關鍵字的條件式格式（使用公式）
    Dim fc As FormatCondition
    Dim firstCellAddr As String
    Dim formulaStr As String
    firstCellAddr = targetRange.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    formulaStr = "=ISNUMBER(SEARCH(" & Chr(34) & keyword & Chr(34) & "," & firstCellAddr & "))"

    Set fc = targetRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:=formulaStr)

    With fc.Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 255, 0) ' 黃色背景
        .TintAndShade = 0
    End With

    With fc.Font
        .Bold = True
        .Color = RGB(192, 0, 0) ' 深紅字
    End With

    MsgBox "已對範圍 " & targetRange.Address & " 套用關鍵字「" & keyword & "」醒目標示格式。", _
        vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "套用條件格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
