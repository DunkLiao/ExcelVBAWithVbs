Option Explicit
'*************************************************************************************
'模組名稱: ClearSortFilterFormatting
'功能說明: 清除工作表上的排序設定與自動篩選，還原為未篩選的原始資料狀態
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub ClearSortAndFilterOnActiveSheet()
    ' 清除作用中工作表的篩選與排序設定
    Dim ws As Worksheet

    On Error GoTo ErrHandler

    Set ws = ActiveSheet

    ' 還原篩選並關閉自動篩選
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
        ws.AutoFilterMode = False
    End If

    ' 清除排序欄位設定
    If ws.Sort.SortFields.Count > 0 Then
        ws.Sort.SortFields.Clear
    End If

    MsgBox "已清除「" & ws.Name & "」工作表的排序與篩選設定。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

Sub ClearSortAndFilterAllSheets()
    ' 清除活頁簿中所有工作表的排序與篩選設定
    Dim ws As Worksheet
    Dim clearCount As Long

    On Error GoTo ErrHandler

    clearCount = 0
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Sheets
        If ws.AutoFilterMode Then
            If ws.FilterMode Then
                ws.ShowAllData
            End If
            ws.AutoFilterMode = False
            clearCount = clearCount + 1
        End If
        If ws.Sort.SortFields.Count > 0 Then
            ws.Sort.SortFields.Clear
        End If
    Next ws

    Application.ScreenUpdating = True

    MsgBox "已清除 " & clearCount & " 個工作表的排序與篩選設定。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

Sub ClearFilterOnlyAllSheets()
    ' 只清除篩選條件（還原顯示所有資料），保留自動篩選按鈕
    Dim ws As Worksheet
    Dim clearCount As Long

    On Error GoTo ErrHandler

    clearCount = 0
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Sheets
        If ws.AutoFilterMode Then
            If ws.FilterMode Then
                ws.ShowAllData
                clearCount = clearCount + 1
            End If
        End If
    Next ws

    Application.ScreenUpdating = True

    MsgBox "已在 " & clearCount & " 個工作表中清除篩選，還原全部資料顯示。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
