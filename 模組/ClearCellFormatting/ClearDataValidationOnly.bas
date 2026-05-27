Option Explicit
Attribute VB_Name = "ClearDataValidationOnly"
'*************************************************************************************

'模組名稱: ClearDataValidationOnly

'功能說明: 只清除選取範圍或整張工作表的資料驗證規則，保留其餘格式與數值

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub ClearDataValidationOnly()

    Dim ws As Worksheet

    Dim rngTarget As Range

    Dim choice As Integer



    Set ws = ActiveSheet



    choice = MsgBox("是否只清除選取範圍的資料驗證？" & vbCrLf & _

        "選 [是] 清除選取範圍" & vbCrLf & _

        "選 [否] 清除整張工作表", _

        vbYesNoCancel + vbQuestion, "清除資料驗證")



    Select Case choice

        Case vbYes

            On Error Resume Next

            Set rngTarget = Selection

            On Error GoTo 0

            If rngTarget Is Nothing Then

                MsgBox "未選取範圍，操作取消。", vbExclamation, "提示"

                Exit Sub

            End If



        Case vbNo

            Set rngTarget = ws.UsedRange



        Case vbCancel

            Exit Sub

    End Select



    ' 清除資料驗證

    On Error Resume Next

    rngTarget.Validation.Delete

    Dim errNum As Long

    errNum = Err.Number

    On Error GoTo 0



    If errNum <> 0 Then

        MsgBox "清除資料驗證時發生錯誤，部分儲存格可能未清除成功。", vbExclamation, "提示"

    Else

        MsgBox "資料驗證已成功清除！" & vbCrLf & "範圍：" & rngTarget.Address, _

            vbInformation, "完成"

    End If

End Sub

