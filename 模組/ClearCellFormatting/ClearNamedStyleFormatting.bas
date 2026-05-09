Attribute VB_Name = "ClearNamedStyleFormatting"
Option Explicit

'*************************************************************************************
'模組名稱: ClearNamedStyleFormatting
'功能說明: 將套用命名樣式的儲存格重設回 Normal 樣式，並清除多餘格式
'
'作者版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 將選取範圍的樣式重設為 Normal
Sub ClearNamedStyleToNormal()
    Dim targetRange As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要重設樣式的儲存格範圍。", vbExclamation, "警告"
        Exit Sub
    End If

    Set targetRange = Selection

    On Error GoTo ErrHandler

    targetRange.Style = "Normal"
    targetRange.ClearFormats

    MsgBox "已將選取範圍重設為 Normal 樣式並清除格式（共 " & _
           targetRange.Cells.Count & " 個儲存格）。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "重設樣式時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 掃描整張工作表，將所有非 Normal 樣式的儲存格重設
Sub ResetAllNonNormalStyles()
    Dim ws       As Worksheet
    Dim cell     As Range
    Dim intCount As Integer

    Set ws = ActiveSheet
    intCount = 0

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    For Each cell In ws.UsedRange
        If cell.Style <> "Normal" Then
            cell.Style = "Normal"
            cell.ClearFormats
            intCount = intCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "已將 " & intCount & " 個非 Normal 樣式的儲存格重設。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "重設樣式時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 列出目前工作簿中所有自訂樣式名稱
Sub ListAllStyles()
    Dim sty     As Style
    Dim strList As String

    strList = "目前活頁簿樣式清單：" & vbCrLf

    For Each sty In ThisWorkbook.Styles
        If Not sty.BuiltIn Then
            strList = strList & "  [自訂] " & sty.Name & vbCrLf
        End If
    Next sty

    If strList = "目前活頁簿樣式清單：" & vbCrLf Then
        strList = strList & "  （無自訂樣式）"
    End If

    MsgBox strList, vbInformation, "樣式清單"
End Sub