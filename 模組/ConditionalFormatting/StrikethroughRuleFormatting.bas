Option Explicit
Attribute VB_Name = "StrikethroughRuleFormatting"
'*************************************************************************************
'模組名稱: StrikethroughRuleFormatting
'功能說明: 依條件套用刪除線格式，當狀態欄位符合指定值時對整列套用刪除線樣式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub ApplyStrikethroughRuleFormatting()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cfRule As FormatCondition
    Dim cfRule2 As FormatCondition
    Dim dataRange As Range

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "工作表內沒有足夠的資料。", vbExclamation, "提示"
        Exit Sub
    End If

    Set dataRange = ws.Range("A2:C" & lastRow)
    dataRange.FormatConditions.Delete

    Set cfRule = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=$C2=""已取消""")

    With cfRule.Font
        .Strikethrough = True
        .Color = RGB(150, 150, 150)
    End With
    cfRule.Interior.Color = RGB(240, 240, 240)

    Set cfRule2 = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=$C2=""已完成""")

    With cfRule2.Font
        .Strikethrough = True
        .Color = RGB(0, 128, 0)
    End With

    MsgBox "已套用刪除線條件式格式規則至 A2:C" & lastRow & "。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "套用刪除線格式失敗"
End Sub