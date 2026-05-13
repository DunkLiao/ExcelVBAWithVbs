Attribute VB_Name = "ExpandAllGroupPivot"
Option Explicit
'*************************************************************************************
'模組名稱: ExpandAllGroupPivot
'功能說明: 展開或折疊樞紐分析表中的所有列群組
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

' 展開樞紐分析表的所有列群組
Sub ExpandAllPivotGroups()
    Dim ws  As Worksheet
    Dim pt  As PivotTable
    Dim pf  As PivotField
    Dim pi  As PivotItem

    Set ws = ActiveSheet

    If ws.PivotTables.Count = 0 Then
        MsgBox "目前工作表沒有樞紐分析表。", vbExclamation, "提示"
        Exit Sub
    End If

    Set pt = ws.PivotTables(1)

    Application.ScreenUpdating = False

    For Each pf In pt.RowFields
        If pf.Name <> "Data" Then
            For Each pi In pf.PivotItems
                On Error Resume Next
                pi.ShowDetail = True
                On Error GoTo 0
            Next pi
        End If
    Next pf

    Application.ScreenUpdating = True
    MsgBox "已展開所有群組。", vbInformation, "完成"
End Sub

' 折疊樞紐分析表的所有列群組
Sub CollapseAllPivotGroups()
    Dim ws  As Worksheet
    Dim pt  As PivotTable
    Dim pf  As PivotField
    Dim pi  As PivotItem

    Set ws = ActiveSheet

    If ws.PivotTables.Count = 0 Then
        MsgBox "目前工作表沒有樞紐分析表。", vbExclamation, "提示"
        Exit Sub
    End If

    Set pt = ws.PivotTables(1)

    Application.ScreenUpdating = False

    For Each pf In pt.RowFields
        If pf.Name <> "Data" Then
            For Each pi In pf.PivotItems
                On Error Resume Next
                pi.ShowDetail = False
                On Error GoTo 0
            Next pi
        End If
    Next pf

    Application.ScreenUpdating = True
    MsgBox "已折疊所有群組。", vbInformation, "完成"
End Sub
