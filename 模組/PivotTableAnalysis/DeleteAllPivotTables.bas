Attribute VB_Name = "DeleteAllPivotTables"
Option Explicit
'*************************************************************************************
'模組名稱: DeleteAllPivotTables
'功能說明: 刪除活頁簿中所有工作表的樞紐分析表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub DeleteAllPivotTables()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim deleteCount As Integer
    Dim ptCount As Integer
    Dim i As Integer
    Dim confirm As Integer

    deleteCount = 0

    confirm = MsgBox("確定要刪除所有工作表中的樞紐分析表嗎？", vbYesNo + vbQuestion, "確認刪除")
    If confirm = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each ws In ThisWorkbook.Worksheets
        ptCount = ws.PivotTables.Count
        For i = ptCount To 1 Step -1
            Set pt = ws.PivotTables(i)
            pt.TableRange2.Clear
            deleteCount = deleteCount + 1
        Next i
    Next ws

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "已刪除 " & deleteCount & " 個樞紐分析表。", vbInformation, "完成"
End Sub

Sub DeletePivotTablesInActiveSheet()
    Dim ws As Worksheet
    Dim ptCount As Integer
    Dim deleteCount As Integer
    Dim i As Integer
    Dim confirm As Integer

    Set ws = ActiveSheet
    ptCount = ws.PivotTables.Count
    deleteCount = 0

    If ptCount = 0 Then
        MsgBox "目前工作表沒有樞紐分析表。", vbInformation, "提示"
        Exit Sub
    End If

    confirm = MsgBox("確定要刪除目前工作表的 " & ptCount & " 個樞紐分析表嗎？", vbYesNo + vbQuestion, "確認刪除")
    If confirm = vbNo Then Exit Sub

    Application.DisplayAlerts = False

    For i = ptCount To 1 Step -1
        ws.PivotTables(i).TableRange2.Clear
        deleteCount = deleteCount + 1
    Next i

    Application.DisplayAlerts = True

    MsgBox "已刪除 " & deleteCount & " 個樞紐分析表。", vbInformation, "完成"
End Sub
