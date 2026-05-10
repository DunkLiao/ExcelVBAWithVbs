'*************************************************************************************
'模組名稱: ChangeDataSourcePivot
'功能說明: 以 VBA 動態變更樞紐分析表的資料來源範圍
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************
Option Explicit

Sub ChangeDataSourcePivot()
    Dim ws          As Worksheet
    Dim pt          As PivotTable
    Dim pc          As PivotCache
    Dim newSource   As String
    Dim srcSheet    As String
    Dim srcRange    As String

    ' 取得使用者指定的樞紐分析表資料來源
    srcSheet = InputBox("請輸入資料來源工作表名稱：", "變更資料來源")
    If srcSheet = "" Then Exit Sub

    srcRange = InputBox("請輸入資料範圍（例如 A1:F100）：", "變更資料來源")
    If srcRange = "" Then Exit Sub

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets(srcSheet)
    newSource = "'" & srcSheet & "'!" & srcRange

    ' 對所有樞紐分析表更新資料來源
    Dim ptSheet     As Worksheet
    Dim updCount    As Integer
    updCount = 0

    For Each ptSheet In ThisWorkbook.Sheets
        For Each pt In ptSheet.PivotTables
            Set pc = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=newSource)
            pt.ChangePivotCache pc
            pt.RefreshTable
            updCount = updCount + 1
        Next pt
    Next ptSheet

    If updCount = 0 Then
        MsgBox "未找到任何樞紐分析表。", vbExclamation, "提示"
    Else
        MsgBox "已成功更新 " & updCount & " 個樞紐分析表的資料來源。", vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
