Option Explicit
Attribute VB_Name = "MergeExcelWithSort"
'*************************************************************************************

'模組名稱: MergeExcelWithSort

'功能說明: 合併指定資料夾內所有 Excel 檔案至主工作表，並依第一欄排序結果

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub MergeExcelWithSort()

    Dim folderPath As String

    Dim fileName As String

    Dim wbSrc As Workbook

    Dim wsDst As Worksheet

    Dim wsData As Worksheet

    Dim dstRow As Long

    Dim srcLastRow As Long

    Dim srcLastCol As Long

    Dim headerCopied As Boolean



    folderPath = InputBox("請輸入要合併的 Excel 資料夾路徑：", "合併並排序", "C:\Data\")

    If folderPath = "" Then Exit Sub

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"



    If Dir(folderPath, vbDirectory) = "" Then

        MsgBox "資料夾不存在：" & folderPath, vbExclamation, "錯誤"

        Exit Sub

    End If



    ' 建立輸出工作表

    Set wsDst = ThisWorkbook.Worksheets.Add

    wsDst.Name = "MergedSorted"

    dstRow = 1

    headerCopied = False



    Application.ScreenUpdating = False



    On Error GoTo ErrHandler



    fileName = Dir(folderPath & "*.xls*")

    Do While fileName <> ""

        If InStr(fileName, "~$") = 0 Then

            Set wbSrc = Workbooks.Open(folderPath & fileName, ReadOnly:=True)

            Set wsData = wbSrc.Worksheets(1)



            srcLastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

            srcLastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column



            If srcLastRow >= 1 Then

                If Not headerCopied Then

                    wsData.Rows(1).Copy wsDst.Rows(dstRow)

                    dstRow = dstRow + 1

                    headerCopied = True

                End If



                If srcLastRow >= 2 Then

                    wsData.Range(wsData.Cells(2, 1), wsData.Cells(srcLastRow, srcLastCol)).Copy _

                        wsDst.Cells(dstRow, 1)

                    dstRow = dstRow + srcLastRow - 1

                End If

            End If



            wbSrc.Close SaveChanges:=False

        End If

        fileName = Dir()

    Loop



    ' 排序合併結果（依第一欄遞增）

    If dstRow > 2 Then

        Dim rngSort As Range

        Set rngSort = wsDst.Range(wsDst.Cells(1, 1), wsDst.Cells(dstRow - 1, _

            wsDst.Cells(1, wsDst.Columns.Count).End(xlToLeft).Column))

        rngSort.Sort Key1:=wsDst.Columns(1), Order1:=xlAscending, Header:=xlYes

    End If



    wsDst.Columns.AutoFit

    Application.ScreenUpdating = True

    MsgBox "合併並排序完成！共 " & (dstRow - 2) & " 筆資料。", vbInformation, "完成"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"

End Sub

