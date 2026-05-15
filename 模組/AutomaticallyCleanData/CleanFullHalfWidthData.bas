Attribute VB_Name = "CleanFullHalfWidthData"

Option Explicit

'*************************************************************************************

'模組名稱: CleanFullHalfWidthData

'功能說明: 將選取範圍中的全形英數與標點轉換為半形字元

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/15

'

'*************************************************************************************



Public Sub RunCleanFullHalfWidthData()

    On Error GoTo ErrorHandler



    Dim targetRange As Range

    Dim cell As Range

    Dim processedCount As Long



    If TypeName(Selection) <> "Range" Then

        MsgBox "請先選取要清理的儲存格範圍。", vbExclamation, "提示"

        Exit Sub

    End If



    Set targetRange = Selection



    For Each cell In targetRange.Cells

        If Not IsError(cell.Value) Then

            If Len(CStr(cell.Value)) > 0 Then

                cell.Value = ConvertFullWidthToHalfWidth(CStr(cell.Value))

                processedCount = processedCount + 1

            End If

        End If

    Next cell



    MsgBox "已完成 " & processedCount & " 個儲存格的全形半形轉換。", vbInformation, "完成"

    Exit Sub



ErrorHandler:

    MsgBox "清理全形半形字元時發生錯誤: " & Err.Description, vbExclamation, "錯誤"

End Sub



Private Function ConvertFullWidthToHalfWidth(ByVal sourceText As String) As String

    Dim i As Long

    Dim oneChar As String

    Dim codePoint As Long

    Dim resultText As String



    For i = 1 To Len(sourceText)

        oneChar = Mid$(sourceText, i, 1)

        codePoint = AscW(oneChar)

        If codePoint < 0 Then codePoint = codePoint + 65536



        Select Case codePoint

            Case 12288

                resultText = resultText & " "

            Case 65281 To 65374

                resultText = resultText & ChrW(codePoint - 65248)

            Case Else

                resultText = resultText & oneChar

        End Select

    Next i



    ConvertFullWidthToHalfWidth = resultText

End Function

