Attribute VB_Name = "ClearDropdownFormatting"
Option Explicit

'*************************************************************************************
'模組名稱: ClearDropdownFormatting
'功能說明: 清除工作表中所有儲存格的下拉式選單資料驗證
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub ClearDropdownFormatting()
    Dim ws As Worksheet
    Dim rngValidation As Range

    Set ws = ThisWorkbook.Worksheets(1)

    On Error Resume Next
    Set rngValidation = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    On Error GoTo 0

    If rngValidation Is Nothing Then
        MsgBox "此工作表沒有資料驗證（下拉選單）設定。", vbInformation
        Exit Sub
    End If

    rngValidation.Validation.Delete

    MsgBox "已清除 " & rngValidation.Count & _
        " 個儲存格的下拉選單設定！", vbInformation
End Sub
