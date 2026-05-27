Option Explicit
Attribute VB_Name = "FilterByCustomCondition"
'*************************************************************************************

'模組名稱: FilterByCustomCondition

'功能說明: 依使用者自訂條件（欄號、運算子、條件值）篩選資料並複製至新工作表

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub FilterByCustomCondition()

    Dim ws As Worksheet

    Dim wsResult As Worksheet

    Dim lastRow As Long

    Dim lastCol As Integer

    Dim i As Long

    Dim colIndex As Integer

    Dim filterOperator As String

    Dim condValue As String

    Dim cellVal As Variant

    Dim matched As Boolean

    Dim dstRow As Long

    Dim colInput As String

    Dim opInput As String



    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column



    If lastRow < 2 Then

        MsgBox "工作表資料不足。", vbExclamation, "提示"

        Exit Sub

    End If



    colInput = InputBox("請輸入要篩選的欄號（例如：2 代表 B 欄）：", "自訂條件篩選", "1")

    If colInput = "" Then Exit Sub

    colIndex = CInt(colInput)

    If colIndex < 1 Or colIndex > lastCol Then

        MsgBox "欄號超出範圍。", vbExclamation, "錯誤"

        Exit Sub

    End If



    opInput = InputBox("請輸入比較運算子（=  <>  >  <  >=  <=  contains）：", "運算子", "=")

    If opInput = "" Then Exit Sub

    filterOperator = Trim(opInput)



    condValue = InputBox("請輸入條件值：", "條件值")

    If condValue = "" Then

        If MsgBox("條件值為空，是否繼續篩選空白儲存格？", vbYesNo + vbQuestion, "提示") = vbNo Then

            Exit Sub

        End If

    End If



    ' 建立結果工作表

    On Error Resume Next

    Application.DisplayAlerts = False

    ThisWorkbook.Worksheets("CustomFilter").Delete

    Application.DisplayAlerts = True

    On Error GoTo 0



    Set wsResult = ThisWorkbook.Worksheets.Add( _

        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

    wsResult.Name = "CustomFilter"



    ' 複製標題

    ws.Rows(1).Copy wsResult.Rows(1)

    dstRow = 2



    Application.ScreenUpdating = False



    For i = 2 To lastRow

        cellVal = ws.Cells(i, colIndex).Value

        matched = False



        Select Case LCase(filterOperator)

            Case "="

                matched = (CStr(cellVal) = condValue)

            Case "<>"

                matched = (CStr(cellVal) <> condValue)

            Case ">"

                If IsNumeric(cellVal) And IsNumeric(condValue) Then

                    matched = (CDbl(cellVal) > CDbl(condValue))

                End If

            Case "<"

                If IsNumeric(cellVal) And IsNumeric(condValue) Then

                    matched = (CDbl(cellVal) < CDbl(condValue))

                End If

            Case ">="

                If IsNumeric(cellVal) And IsNumeric(condValue) Then

                    matched = (CDbl(cellVal) >= CDbl(condValue))

                End If

            Case "<="

                If IsNumeric(cellVal) And IsNumeric(condValue) Then

                    matched = (CDbl(cellVal) <= CDbl(condValue))

                End If

            Case "contains"

                matched = (InStr(1, CStr(cellVal), condValue, vbTextCompare) > 0)

            Case Else

                MsgBox "不支援的運算子：" & filterOperator, vbExclamation, "錯誤"

                Application.ScreenUpdating = True

                Application.DisplayAlerts = False

                wsResult.Delete

                Application.DisplayAlerts = True

                Exit Sub

        End Select



        If matched Then

            ws.Rows(i).Copy wsResult.Rows(dstRow)

            dstRow = dstRow + 1

        End If

    Next i



    wsResult.Columns.AutoFit

    Application.ScreenUpdating = True



    MsgBox "自訂條件篩選完成！共篩出 " & (dstRow - 2) & " 筆資料，結果在工作表：" & wsResult.Name, _

        vbInformation, "完成"

End Sub

