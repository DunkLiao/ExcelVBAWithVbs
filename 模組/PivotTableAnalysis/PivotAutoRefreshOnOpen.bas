Option Explicit
'*************************************************************************************
'模組名稱: PivotAutoRefreshOnOpen
'功能說明: 設定活頁簿中所有樞紐分析表於開啟時自動更新，並提供手動全部重新整理功能
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub EnablePivotAutoRefreshOnOpen()
    ' 設定所有樞紐分析表在開啟活頁簿時自動重新整理
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pivotCount As Long

    On Error GoTo ErrHandler

    pivotCount = 0

    For Each ws In ThisWorkbook.Sheets
        For Each pt In ws.PivotTables
            pt.RefreshOnFileOpen = True
            pivotCount = pivotCount + 1
        Next pt
    Next ws

    If pivotCount = 0 Then
        MsgBox "本活頁簿中找不到任何樞紐分析表！", vbExclamation, "提示"
    Else
        MsgBox "已設定 " & pivotCount & " 個樞紐分析表於開啟時自動重新整理。", _
               vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

Sub DisablePivotAutoRefreshOnOpen()
    ' 停用所有樞紐分析表的自動重新整理設定
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pivotCount As Long

    On Error GoTo ErrHandler

    pivotCount = 0

    For Each ws In ThisWorkbook.Sheets
        For Each pt In ws.PivotTables
            pt.RefreshOnFileOpen = False
            pivotCount = pivotCount + 1
        Next pt
    Next ws

    MsgBox "已停用 " & pivotCount & " 個樞紐分析表的自動重新整理設定。", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

Sub RefreshAllPivotTablesNow()
    ' 立即重新整理活頁簿中所有樞紐分析表
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pivotCount As Long

    On Error GoTo ErrHandler

    pivotCount = 0
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Sheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
            pivotCount = pivotCount + 1
        Next pt
    Next ws

    Application.ScreenUpdating = True

    If pivotCount = 0 Then
        MsgBox "本活頁簿中找不到任何樞紐分析表！", vbExclamation, "提示"
    Else
        MsgBox "已重新整理 " & pivotCount & " 個樞紐分析表。", vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
