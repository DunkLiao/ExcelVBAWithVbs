Attribute VB_Name = "ClearPivotTableFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearPivotTableFormatting
'功能說明: 清除工作表中所有樞紐分析表的格式設定，還原為預設外觀
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestClearPivotTableFormatting()
    Call ClearAllPivotFormats
End Sub

' 清除所有樞紐分析表格式
Sub ClearAllPivotFormats()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim intCount As Integer

    On Error GoTo ErrHandler
    intCount = 0

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            Call ResetSinglePivotFormat(pt)
            intCount = intCount + 1
        Next pt
    Next ws

    If intCount = 0 Then
        MsgBox "目前活頁簿中找不到任何樞紐分析表。", vbInformation, "提示"
    Else
        MsgBox "已清除 " & intCount & " 個樞紐分析表的格式設定！", _
            vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 重設單一樞紐分析表格式
Private Sub ResetSinglePivotFormat(ByVal pt As PivotTable)
    Dim pf As PivotField
    On Error Resume Next
    pt.TableStyle2 = ""
    pt.ShowTableStyleRowStripes = False
    pt.ShowTableStyleColumnStripes = False
    pt.ShowTableStyleColumnHeaders = True
    For Each pf In pt.DataFields
        pf.NumberFormat = "General"
    Next pf
    On Error GoTo 0
End Sub

' 清除目前工作表上的樞紐分析表格式
Sub ClearActivePivotFormat()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim intCount As Integer

    On Error GoTo ErrHandler
    Set ws = ActiveSheet
    intCount = 0

    For Each pt In ws.PivotTables
        Call ResetSinglePivotFormat(pt)
        intCount = intCount + 1
    Next pt

    If intCount = 0 Then
        MsgBox "目前工作表找不到任何樞紐分析表。", vbInformation, "提示"
    Else
        MsgBox "已清除目前工作表中 " & intCount & " 個樞紐分析表的格式！", _
            vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
