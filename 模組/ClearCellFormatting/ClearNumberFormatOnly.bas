Attribute VB_Name = "ClearNumberFormatOnly"
Option Explicit

'*************************************************************************************
'模組名稱: ClearNumberFormatOnly
'功能說明: 只清除數字格式，將選取範圍重設為 General，保留其他格式
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 清除選取範圍的數字格式
Sub ClearNumberFormatOnly()
    Dim targetRange As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除數字格式的儲存格範圍。", vbExclamation, "警告"
        Exit Sub
    End If

    Set targetRange = Selection

    On Error GoTo ErrHandler

    targetRange.NumberFormat = "General"

    MsgBox "已將選取範圍的數字格式重設為 General（共 " & _
           targetRange.Cells.Count & " 個儲存格）。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除數字格式時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除整張工作表使用範圍的數字格式
Sub ClearNumberFormatInSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    On Error GoTo ErrHandler

    ws.UsedRange.NumberFormat = "General"

    MsgBox "已清除工作表「" & ws.Name & "」所有使用範圍的數字格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除數字格式時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 示範各種數字格式的套用與清除
Sub DemoNumberFormatClear()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ws.Range("A1").Value = 1234567.89
    ws.Range("A2").Value = 0.75
    ws.Range("A3").Value = 45000
    ws.Range("A4").Value = -500

    ws.Range("A1").NumberFormat = "#,##0.00"
    ws.Range("A2").NumberFormat = "0.00%"
    ws.Range("A3").NumberFormat = "yyyy/mm/dd"
    ws.Range("A4").NumberFormat = "$#,##0.00;[Red]($#,##0.00)"

    MsgBox "已套用各種數字格式於 A1:A4，請選取後執行 ClearNumberFormatOnly。", _
           vbInformation, "提示"
End Sub