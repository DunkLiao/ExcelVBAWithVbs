'*************************************************************************************
'模組名稱: TextRuleFormatting
'功能說明: 以 VBA 建立以文字條件為基礎的條件式格式設定範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************
Option Explicit

Sub ApplyTextRuleFormatting()
    Dim ws          As Worksheet
    Dim rng         As Range
    Dim fc          As FormatCondition

    Set ws = ActiveSheet
    ' 使用目前工作表的 A 欄資料
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "A 欄資料不足，請先輸入資料再執行。", vbExclamation, "提示"
        Exit Sub
    End If

    Set rng = ws.Range("A2:A" & lastRow)

    ' 清除舊有條件式格式
    rng.FormatConditions.Delete

    ' 規則 1：包含「通過」→ 綠底白字
    Set fc = rng.FormatConditions.Add( _
        Type:=xlTextString, _
        String:="通過", _
        TextOperator:=xlContains)
    With fc.Interior
        .Color = RGB(0, 176, 80)
    End With
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.Bold = True

    ' 規則 2：包含「不通過」→ 紅底白字
    Set fc = rng.FormatConditions.Add( _
        Type:=xlTextString, _
        String:="不通過", _
        TextOperator:=xlContains)
    With fc.Interior
        .Color = RGB(255, 0, 0)
    End With
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.Bold = True

    ' 規則 3：包含「待審」→ 黃底黑字
    Set fc = rng.FormatConditions.Add( _
        Type:=xlTextString, _
        String:="待審", _
        TextOperator:=xlContains)
    With fc.Interior
        .Color = RGB(255, 255, 0)
    End With
    fc.Font.Color = RGB(0, 0, 0)

    MsgBox "文字條件式格式設定完成！", vbInformation, "完成"
End Sub
