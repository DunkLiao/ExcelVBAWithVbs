Option Explicit
Attribute VB_Name = "BatchConvertFormulas"
'*************************************************************************************

'模組名稱: BatchConvertFormulas

'功能說明: 批次將選取範圍內的公式轉換為靜態數值（或反向將數值轉換為文字）

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub BatchConvertFormulas()

    Dim rngTarget As Range

    Dim cell As Range

    Dim choice As Integer



    On Error Resume Next

    Set rngTarget = Selection

    On Error GoTo 0



    If rngTarget Is Nothing Then

        MsgBox "請先選取要轉換的範圍。", vbExclamation, "提示"

        Exit Sub

    End If



    choice = MsgBox("選擇轉換方式：" & vbCrLf & _

        "[是] 公式 -> 靜態數值（貼上值）" & vbCrLf & _

        "[否] 數值 -> 文字格式" & vbCrLf & _

        "[取消] 結束", _

        vbYesNoCancel + vbQuestion, "批次轉換公式")



    Application.ScreenUpdating = False



    Select Case choice

        Case vbYes

            ' 公式轉靜態數值

            Dim formulaCount As Long

            formulaCount = 0

            For Each cell In rngTarget.Cells

                If cell.HasFormula Then

                    cell.Value = cell.Value

                    formulaCount = formulaCount + 1

                End If

            Next cell

            Application.ScreenUpdating = True

            MsgBox "已將 " & formulaCount & " 個公式轉換為靜態數值。", vbInformation, "完成"



        Case vbNo

            ' 數值轉文字格式

            Dim numCount As Long

            numCount = 0

            Dim numVal As Variant

            For Each cell In rngTarget.Cells

                If Not cell.HasFormula And IsNumeric(cell.Value) And cell.Value <> "" Then

                    numVal = cell.Value

                    cell.NumberFormat = "@"

                    cell.Value = CStr(numVal)

                    numCount = numCount + 1

                End If

            Next cell

            Application.ScreenUpdating = True

            MsgBox "已將 " & numCount & " 個數值轉換為文字格式。", vbInformation, "完成"



        Case vbCancel

            Application.ScreenUpdating = True

            Exit Sub

    End Select

End Sub

