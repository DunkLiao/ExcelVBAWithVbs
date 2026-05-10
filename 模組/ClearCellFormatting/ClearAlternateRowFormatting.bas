Attribute VB_Name = "ClearAlternateRowFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearAlternateRowFormatting
'功能說明: 清除目前工作表中指定範圍內奇數列或偶數列的儲存格格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點：清除奇數列格式
Sub TestClearOddRowFormatting()
    Call ClearAlternateRowFormatting(True)
End Sub

' 範例進入點：清除偶數列格式
Sub TestClearEvenRowFormatting()
    Call ClearAlternateRowFormatting(False)
End Sub

' 清除目前選取範圍中的交替列格式
' clearOddRows: True = 清除奇數列，False = 清除偶數列
Sub ClearAlternateRowFormatting(ByVal clearOddRows As Boolean)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim i As Long
    Dim rowType As String

    Set ws = ActiveSheet

    If TypeName(Selection) = "Range" And Selection.Cells.Count > 1 Then
        Set targetRange = Selection
    Else
        Set targetRange = ws.UsedRange
    End If

    firstRow = targetRange.Row
    lastRow = firstRow + targetRange.Rows.Count - 1
    firstCol = targetRange.Column
    lastCol = firstCol + targetRange.Columns.Count - 1

    If clearOddRows Then
        rowType = "奇數"
    Else
        rowType = "偶數"
    End If

    Dim clearedCount As Long
    clearedCount = 0

    Application.ScreenUpdating = False

    For i = firstRow To lastRow
        Dim isOdd As Boolean
        isOdd = (i Mod 2 = 1)

        If (clearOddRows And isOdd) Or (Not clearOddRows And Not isOdd) Then
            With ws.Range(ws.Cells(i, firstCol), ws.Cells(i, lastCol))
                .Interior.ColorIndex = xlNone
                .Font.Bold = False
                .Font.Italic = False
                .Font.ColorIndex = xlAutomatic
                .Borders.LineStyle = xlNone
                .NumberFormat = "General"
            End With
            clearedCount = clearedCount + 1
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "已清除 " & clearedCount & " 個" & rowType & "列的格式。", _
           vbInformation, "清除完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清除格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
