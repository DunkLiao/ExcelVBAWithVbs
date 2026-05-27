Option Explicit
Attribute VB_Name = "CleanTimeFormat"
'*************************************************************************************

'模組名稱: CleanTimeFormat

'功能說明: 自動清理工作表中格式不一致的時間資料，統一轉換為 HH:MM:SS 格式

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub CleanTimeFormat()

    Dim ws As Worksheet

    Dim rngTarget As Range

    Dim cell As Range

    Dim rawVal As String

    Dim fixedCount As Long

    Dim skipCount As Long

    Dim choice As Integer

    Dim hh As Integer

    Dim mm As Integer



    Set ws = ActiveSheet



    choice = MsgBox("是否只清理選取範圍？" & vbCrLf & _

        "[是] 清理選取範圍" & vbCrLf & _

        "[否] 清理整張工作表已使用範圍", _

        vbYesNo + vbQuestion, "清理時間格式")



    If choice = vbYes Then

        On Error Resume Next

        Set rngTarget = Selection

        On Error GoTo 0

        If rngTarget Is Nothing Then

            MsgBox "未選取範圍，操作取消。", vbExclamation, "提示"

            Exit Sub

        End If

    Else

        Set rngTarget = ws.UsedRange

    End If



    Application.ScreenUpdating = False



    fixedCount = 0

    skipCount = 0



    For Each cell In rngTarget.Cells

        rawVal = Trim(CStr(cell.Value))



        If rawVal <> "" Then

            Dim timeVal As Variant

            timeVal = Empty



            ' 格式：HH:MM 或 HH:MM:SS

            If rawVal Like "##:##" Or rawVal Like "##:##:##" Then

                On Error Resume Next

                timeVal = CDate(rawVal)

                On Error GoTo 0



            ' 格式：HHMM（四位數）

            ElseIf Len(rawVal) = 4 And IsNumeric(rawVal) Then

                hh = CInt(Left(rawVal, 2))

                mm = CInt(Right(rawVal, 2))

                If hh >= 0 And hh <= 23 And mm >= 0 And mm <= 59 Then

                    On Error Resume Next

                    timeVal = TimeSerial(hh, mm, 0)

                    On Error GoTo 0

                End If



            ' 格式：H:MM 或含 AM/PM

            ElseIf InStr(rawVal, ":") > 0 Then

                On Error Resume Next

                timeVal = CDate(rawVal)

                On Error GoTo 0

            End If



            ' 若成功解析，統一套用 HH:MM:SS 格式

            If Not IsEmpty(timeVal) Then

                cell.Value = timeVal

                cell.NumberFormat = "HH:MM:SS"

                fixedCount = fixedCount + 1

            Else

                skipCount = skipCount + 1

            End If

        End If

    Next cell



    Application.ScreenUpdating = True



    MsgBox "時間格式清理完成！" & vbCrLf & _

        "已修正：" & fixedCount & " 個儲存格" & vbCrLf & _

        "略過（無法識別）：" & skipCount & " 個儲存格", vbInformation, "完成"

End Sub

