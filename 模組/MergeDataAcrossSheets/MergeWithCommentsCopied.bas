Option Explicit
Attribute VB_Name = "MergeWithCommentsCopied"
'*************************************************************************************

'模組名稱: MergeWithCommentsCopied

'功能說明: 跨工作表合併資料時，保留原始儲存格批注（Comment）至目標工作表

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub MergeWithCommentsCopied()

    Dim wsSummary As Worksheet

    Dim wsSource As Worksheet

    Dim dstRow As Long

    Dim srcLastRow As Long

    Dim srcLastCol As Long

    Dim i As Long

    Dim j As Integer

    Dim commentText As String

    Dim srcCell As Range

    Dim dstCell As Range

    Dim headerCopied As Boolean



    ' 建立彙整工作表

    On Error Resume Next

    Application.DisplayAlerts = False

    ThisWorkbook.Worksheets("MergedWithComments").Delete

    Application.DisplayAlerts = True

    On Error GoTo 0



    Set wsSummary = ThisWorkbook.Worksheets.Add( _

        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

    wsSummary.Name = "MergedWithComments"



    dstRow = 1

    headerCopied = False



    Application.ScreenUpdating = False



    ' 遍歷所有工作表（排除彙整表自身）

    For Each wsSource In ThisWorkbook.Worksheets

        If wsSource.Name <> wsSummary.Name Then

            srcLastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

            srcLastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column



            If srcLastRow >= 1 And srcLastCol >= 1 Then

                Dim startRow As Long

                If Not headerCopied Then

                    startRow = 1

                    headerCopied = True

                Else

                    startRow = 2

                End If



                For i = startRow To srcLastRow

                    For j = 1 To srcLastCol

                        Set srcCell = wsSource.Cells(i, j)

                        Set dstCell = wsSummary.Cells(dstRow + (i - startRow), j)



                        ' 複製數值

                        dstCell.Value = srcCell.Value



                        ' 複製格式

                        srcCell.Copy

                        dstCell.PasteSpecial Paste:=xlPasteFormats

                        Application.CutCopyMode = False



                        ' 複製批注

                        If Not srcCell.Comment Is Nothing Then

                            commentText = srcCell.Comment.Text

                            dstCell.AddComment commentText

                            dstCell.Comment.Visible = False

                        End If

                    Next j

                Next i



                dstRow = dstRow + (srcLastRow - startRow + 1)

            End If

        End If

    Next wsSource



    wsSummary.Columns.AutoFit

    Application.ScreenUpdating = True



    MsgBox "跨表合併（含批注）完成！彙整至工作表：" & wsSummary.Name, vbInformation, "完成"

End Sub

