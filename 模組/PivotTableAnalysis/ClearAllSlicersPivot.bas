Attribute VB_Name = "ClearAllSlicersPivot"
Option Explicit
'*************************************************************************************
'模組名稱: ClearAllSlicersPivot
'功能說明: 清除活頁簿中所有樞紐分析表的交叉分析篩選器（Slicer）篩選狀態
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestClearAllSlicers()
    Call ClearAllSlicersPivot(ThisWorkbook)
End Sub

Sub ClearAllSlicersPivot(ByVal wb As Workbook)
    Dim sc           As SlicerCache
    Dim clearedCount As Integer

    clearedCount = 0
    Application.ScreenUpdating = False

    For Each sc In wb.SlicerCaches
        On Error Resume Next
        sc.ClearManualFilter
        On Error GoTo 0
        clearedCount = clearedCount + 1
    Next sc

    Application.ScreenUpdating = True

    If clearedCount > 0 Then
        MsgBox "已清除 " & clearedCount & " 個交叉分析篩選器的篩選狀態。", _
               vbInformation, "完成"
    Else
        MsgBox "目前活頁簿中沒有交叉分析篩選器。", vbInformation, "提示"
    End If
End Sub

Sub ClearSlicersOnSheet(ByVal ws As Worksheet)
    Dim slicerObj    As Slicer
    Dim clearedCount As Integer

    clearedCount = 0
    Application.ScreenUpdating = False

    For Each slicerObj In ws.Slicers
        On Error Resume Next
        slicerObj.SlicerCache.ClearManualFilter
        On Error GoTo 0
        clearedCount = clearedCount + 1
    Next slicerObj

    Application.ScreenUpdating = True

    MsgBox "已清除工作表『" & ws.Name & "』中 " & clearedCount & _
           " 個篩選器狀態。", vbInformation, "完成"
End Sub

Sub ListAllSlicers()
    Dim wb  As Workbook
    Dim sc  As SlicerCache
    Dim msg As String

    Set wb = ThisWorkbook
    msg = "活頁簿中的交叉分析篩選器：" & Chr(13)

    If wb.SlicerCaches.Count = 0 Then
        MsgBox "目前沒有任何交叉分析篩選器。", vbInformation, "提示"
        Exit Sub
    End If

    For Each sc In wb.SlicerCaches
        msg = msg & "  - " & sc.Name & _
              "（來源欄位：" & sc.SourceName & "）" & Chr(13)
    Next sc

    MsgBox msg, vbInformation, "篩選器清單"
End Sub
