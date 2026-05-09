Attribute VB_Name = "ClearProtectionFormatting"
Option Explicit

'*************************************************************************************
'模組名稱: ClearProtectionFormatting
'功能說明: 清除儲存格的鎖定與隱藏保護設定，並同時清除格式
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 清除選取範圍的保護設定（鎖定+隱藏）並清除格式
Sub ClearProtectionAndFormatting()
    Dim targetRange As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除保護設定的儲存格範圍。", vbExclamation, "警告"
        Exit Sub
    End If

    Set targetRange = Selection

    On Error GoTo ErrHandler

    ' 先解除工作表保護（若已保護）
    If ActiveSheet.ProtectContents Then
        MsgBox "工作表目前已保護，請先解除保護再執行此操作。", vbExclamation, "警告"
        Exit Sub
    End If

    With targetRange
        .Locked = False
        .FormulaHidden = False
        .ClearFormats
    End With

    MsgBox "已清除選取範圍的鎖定、隱藏保護設定與格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "清除保護設定時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次解鎖整張工作表所有儲存格並清除格式
Sub UnlockAllCellsAndClearFormats()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.ProtectContents Then
        MsgBox "工作表目前已保護，請先解除保護再執行此操作。", vbExclamation, "警告"
        Exit Sub
    End If

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    ws.Cells.Locked = False
    ws.Cells.FormulaHidden = False
    ws.UsedRange.ClearFormats

    Application.ScreenUpdating = True
    MsgBox "已解鎖工作表「" & ws.Name & "」所有儲存格並清除使用範圍格式。", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "操作時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub