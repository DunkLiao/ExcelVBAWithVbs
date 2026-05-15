Attribute VB_Name = "PivotTableFromListObject"

Option Explicit

'*************************************************************************************

'模組名稱: PivotTableFromListObject

'功能說明: 從作用中工作表的結構化表格自動建立樞紐分析表

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/15

'

'*************************************************************************************



Public Sub RunPivotTableFromListObject()

    On Error GoTo ErrorHandler



    Dim wsSource As Worksheet

    Dim wsPivot As Worksheet

    Dim lo As ListObject

    Dim pc As PivotCache

    Dim pt As PivotTable

    Dim rowFieldIndex As Long

    Dim columnFieldIndex As Long

    Dim valueFieldIndex As Long

    Dim valueFunction As XlConsolidationFunction

    Dim pivotSheetName As String



    Set wsSource = ActiveSheet

    If wsSource Is Nothing Then Exit Sub

    If wsSource.ListObjects.Count = 0 Then

        MsgBox "作用中工作表找不到 Excel 表格。", vbExclamation, "提示"

        Exit Sub

    End If



    Set lo = wsSource.ListObjects(1)

    Call ResolvePivotFieldIndexes(lo, rowFieldIndex, columnFieldIndex, valueFieldIndex, valueFunction)



    pivotSheetName = GetUniquePivotSheetName("樞紐分析_" & lo.Name)

    Set wsPivot = ThisWorkbook.Worksheets.Add(After:=wsSource)

    wsPivot.Name = pivotSheetName



    Set pc = ThisWorkbook.PivotCaches.Create( _

        SourceType:=xlDatabase, _

        SourceData:=lo.Range.Address(ReferenceStyle:=xlR1C1, External:=True))



    Set pt = pc.CreatePivotTable( _

        TableDestination:=wsPivot.Range("A3"), _

        TableName:="Pivot_" & Replace(lo.Name, " ", "_"))



    With pt

        .PivotFields(lo.ListColumns(rowFieldIndex).Name).Orientation = xlRowField

        .PivotFields(lo.ListColumns(rowFieldIndex).Name).Position = 1



        If columnFieldIndex > 0 Then

            .PivotFields(lo.ListColumns(columnFieldIndex).Name).Orientation = xlColumnField

            .PivotFields(lo.ListColumns(columnFieldIndex).Name).Position = 1

        End If



        .AddDataField _

            .PivotFields(lo.ListColumns(valueFieldIndex).Name), _

            IIf(valueFunction = xlSum, "總和", "筆數") & lo.ListColumns(valueFieldIndex).Name, _

            valueFunction

        .RowAxisLayout xlTabularRow

    End With



    wsPivot.Columns.AutoFit

    MsgBox "已從結構化表格建立樞紐分析表。", vbInformation, "完成"

    Exit Sub



ErrorHandler:

    MsgBox "建立樞紐分析表時發生錯誤: " & Err.Description, vbExclamation, "錯誤"

End Sub



Private Sub ResolvePivotFieldIndexes(ByVal lo As ListObject, ByRef rowFieldIndex As Long, _

    ByRef columnFieldIndex As Long, ByRef valueFieldIndex As Long, ByRef valueFunction As XlConsolidationFunction)



    Dim i As Long



    rowFieldIndex = 0

    columnFieldIndex = 0

    valueFieldIndex = 0

    valueFunction = xlCount



    For i = 1 To lo.ListColumns.Count

        If IsNumericListColumn(lo.ListColumns(i)) Then

            If valueFieldIndex = 0 Then

                valueFieldIndex = i

                valueFunction = xlSum

            ElseIf columnFieldIndex = 0 Then

                columnFieldIndex = i

            End If

        Else

            If rowFieldIndex = 0 Then

                rowFieldIndex = i

            ElseIf columnFieldIndex = 0 Then

                columnFieldIndex = i

            End If

        End If

    Next i



    If rowFieldIndex = 0 Then rowFieldIndex = 1

    If valueFieldIndex = 0 Then

        valueFieldIndex = rowFieldIndex

        valueFunction = xlCount

    End If

    If columnFieldIndex = rowFieldIndex Or columnFieldIndex = valueFieldIndex Then

        columnFieldIndex = 0

    End If

End Sub



Private Function IsNumericListColumn(ByVal listColumn As ListColumn) As Boolean

    Dim cell As Range



    If listColumn.DataBodyRange Is Nothing Then Exit Function



    For Each cell In listColumn.DataBodyRange.Cells

        If Len(CStr(cell.Value)) > 0 Then

            If IsNumeric(cell.Value) Then

                IsNumericListColumn = True

            End If

            Exit Function

        End If

    Next cell

End Function



Private Function GetUniquePivotSheetName(ByVal baseName As String) As String

    Dim candidate As String

    Dim indexValue As Long



    candidate = Left$(baseName, 31)

    indexValue = 1



    Do While PivotSheetExists(candidate)

        candidate = Left$(baseName, 28) & Format(indexValue, "000")

        indexValue = indexValue + 1

    Loop



    GetUniquePivotSheetName = candidate

End Function



Private Function PivotSheetExists(ByVal sheetName As String) As Boolean

    On Error Resume Next

    PivotSheetExists = Not ws Is Nothing

    On Error GoTo 0

End Function

