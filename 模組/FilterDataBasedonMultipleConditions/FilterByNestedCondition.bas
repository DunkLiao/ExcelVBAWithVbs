Attribute VB_Name = "FilterByNestedCondition"
Option Explicit

'*************************************************************************************
'模組名稱: FilterByNestedCondition
'功能說明: 依巢狀 AND/OR 條件篩選資料並複製至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub FilterByNestedCondition()
    '條件：欄C值 > 100 AND（欄D值 = "Y" OR 欄E值 >= 50）
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim destRow As Long
    Dim colC As Variant
    Dim colD As String
    Dim colE As Variant
    Dim condDE As Boolean

    Set wsSrc = ThisWorkbook.Worksheets(1)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("篩選結果").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDest = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsDest.Name = "篩選結果"

    wsSrc.Rows(1).Copy wsDest.Rows(1)
    destRow = 2

    For i = 2 To lastRow
        colC = wsSrc.Cells(i, 3).Value
        colD = Trim(CStr(wsSrc.Cells(i, 4).Value))
        colE = wsSrc.Cells(i, 5).Value
        condDE = False

        If IsNumeric(colC) Then
            If CDbl(colC) > 100 Then
                If colD = "Y" Then
                    condDE = True
                ElseIf IsNumeric(colE) Then
                    If CDbl(colE) >= 50 Then
                        condDE = True
                    End If
                End If
                If condDE Then
                    wsSrc.Rows(i).Copy wsDest.Rows(destRow)
                    destRow = destRow + 1
                End If
            End If
        End If
    Next i

    wsDest.Columns.AutoFit
    MsgBox "篩選完成，共找到 " & (destRow - 2) & " 筆符合資料！", vbInformation
End Sub
