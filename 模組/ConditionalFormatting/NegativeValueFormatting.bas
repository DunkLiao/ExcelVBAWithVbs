Attribute VB_Name = "NegativeValueFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: NegativeValueFormatting
'功能說明: 自動為選取範圍中的負值儲存格套用紅色背景及白色字體的條件式格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestNegativeValueFormatting()
    Call ApplyNegativeValueFormatting
End Sub

' 為選取範圍套用負值條件式格式
Sub ApplyNegativeValueFormatting()
    On Error GoTo ErrorHandler

    Dim targetRange As Range
    Dim ws As Worksheet
    Dim fc As FormatCondition
    Dim rangeAddress As String

    rangeAddress = InputBox("請輸入要套用負值格式的範圍位址：" & vbCrLf & _
                            "例如：B2:E20", "設定範圍", "B2:E20")
    If rangeAddress = "" Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If

    Set ws = ActiveSheet

    On Error Resume Next
    Set targetRange = ws.Range(rangeAddress)
    On Error GoTo ErrorHandler

    If targetRange Is Nothing Then
        MsgBox "範圍位址無效，請重新輸入", vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 清除既有條件式格式
    targetRange.FormatConditions.Delete

    ' 新增「儲存格值 < 0」的條件式格式
    Set fc = targetRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlLess, _
        Formula1:="0")

    ' 設定格式：紅色背景、白色粗體字
    With fc.Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(192, 0, 0)
        .TintAndShade = 0
    End With

    With fc.Font
        .Color = RGB(255, 255, 255)
        .Bold = True
    End With

    fc.StopIfTrue = False

    MsgBox "已為範圍 " & rangeAddress & " 套用負值條件式格式！" & vbCrLf & _
           "負值將顯示為：紅色背景、白色粗體", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "套用條件式格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
