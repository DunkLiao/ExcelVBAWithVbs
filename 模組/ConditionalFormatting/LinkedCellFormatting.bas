Option Explicit
Attribute VB_Name = "LinkedCellFormatting"
'*************************************************************************************

'模組名稱: LinkedCellFormatting

'功能說明: 依參考儲存格的值，動態套用條件式格式至目標範圍（連結儲存格格式）

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub ApplyLinkedCellFormatting()

    Dim ws As Worksheet

    Dim rngTarget As Range

    Dim rngRef As Range

    Dim refCol As String

    Dim refRow As String

    Dim targetFirstAddr As String



    Set ws = ActiveSheet



    ' 取得目標範圍

    On Error Resume Next

    Set rngTarget = Application.InputBox( _

        "請選取要套用格式的目標範圍：", "連結儲存格格式", Type:=8)

    On Error GoTo 0

    If rngTarget Is Nothing Then Exit Sub



    ' 取得參考儲存格

    On Error Resume Next

    Set rngRef = Application.InputBox( _

        "請選取參考值的儲存格（單一儲存格）：", "選取參考儲存格", Type:=8)

    On Error GoTo 0

    If rngRef Is Nothing Then Exit Sub



    If rngRef.Cells.Count > 1 Then

        MsgBox "請只選取一個儲存格作為參考。", vbExclamation, "錯誤"

        Exit Sub

    End If



    ' 取得參考儲存格的絕對欄列位址

    refCol = "$" & rngRef.Cells(1, 1).Address(True, True)

    targetFirstAddr = rngTarget.Cells(1, 1).Address(False, False)



    ' 清除目標範圍既有條件格式

    rngTarget.FormatConditions.Delete



    ' 套用條件 1：目標值 > 參考值 -> 綠底

    With rngTarget.FormatConditions.Add( _

        Type:=xlExpression, _

        Formula1:="=" & targetFirstAddr & ">" & rngRef.Cells(1, 1).Address(True, True))

        .Interior.Color = RGB(144, 238, 144)

        .Font.Color = RGB(0, 100, 0)

        .Font.Bold = True

    End With



    ' 套用條件 2：目標值 = 參考值 -> 黃底

    With rngTarget.FormatConditions.Add( _

        Type:=xlExpression, _

        Formula1:="=" & targetFirstAddr & "=" & rngRef.Cells(1, 1).Address(True, True))

        .Interior.Color = RGB(255, 255, 153)

        .Font.Color = RGB(153, 102, 0)

        .Font.Bold = True

    End With



    ' 套用條件 3：目標值 < 參考值 -> 紅底

    With rngTarget.FormatConditions.Add( _

        Type:=xlExpression, _

        Formula1:="=" & targetFirstAddr & "<" & rngRef.Cells(1, 1).Address(True, True))

        .Interior.Color = RGB(255, 182, 193)

        .Font.Color = RGB(139, 0, 0)

        .Font.Bold = True

    End With



    MsgBox "連結儲存格格式套用完成！" & vbCrLf & _

        "目標範圍：" & rngTarget.Address & vbCrLf & _

        "參考儲存格：" & rngRef.Address, vbInformation, "完成"

End Sub

