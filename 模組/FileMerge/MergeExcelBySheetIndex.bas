Attribute VB_Name = "MergeExcelBySheetIndex"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelBySheetIndex
'功能說明: 合併指定資料夾中所有Excel檔案的第一張工作表資料至主工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub MergeExcelBySheetIndex()
    Dim folderPath As String
    Dim fileName As String
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim destRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim sheetIndex As Integer
    Dim isFirst As Boolean

    sheetIndex = 1

    folderPath = InputBox("請輸入包含Excel檔案的資料夾路徑：", "選擇資料夾")
    If folderPath = "" Then Exit Sub
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    On Error Resume Next
    Set destWs = ThisWorkbook.Worksheets("合併結果")
    On Error GoTo 0

    If destWs Is Nothing Then
        Set destWs = ThisWorkbook.Worksheets.Add
        destWs.Name = "合併結果"
    Else
        destWs.Cells.Clear
    End If

    destRow = 1
    isFirst = True

    fileName = Dir(folderPath & "*.xlsx")

    Do While fileName <> ""
        Application.ScreenUpdating = False
        Set srcWb = Workbooks.Open(folderPath & fileName, ReadOnly:=True)

        If srcWb.Worksheets.Count >= sheetIndex Then
            Set srcWs = srcWb.Worksheets(sheetIndex)
            lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
            lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

            If isFirst Then
                srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol)).Copy _
                    destWs.Cells(destRow, 1)
                destRow = destRow + lastRow
                isFirst = False
            Else
                If lastRow > 1 Then
                    srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(lastRow, lastCol)).Copy _
                        destWs.Cells(destRow, 1)
                    destRow = destRow + lastRow - 1
                End If
            End If
        End If

        srcWb.Close SaveChanges:=False
        fileName = Dir()
    Loop

    Application.ScreenUpdating = True
    MsgBox "合併完成！共 " & (destRow - 1) & " 列資料。", vbInformation, "完成"
End Sub
