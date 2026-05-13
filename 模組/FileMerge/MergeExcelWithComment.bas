Attribute VB_Name = "MergeExcelWithComment"
Option Explicit
'*************************************************************************************
'模組名稱: 合併含備注的Excel
'功能說明: 合併資料夾內多個Excel檔案，並保留各儲存格的備注(Comments)
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub MergeExcelWithComment()
    Dim folderPath As String
    Dim fileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim destRow As Long
    Dim srcRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Long
    Dim cmt As Comment
    Dim isFirstFile As Boolean
    Dim startRow As Long

    folderPath = InputBox("請輸入要合併的資料夾路徑：", "選擇資料夾", "C:\")
    If folderPath = "" Then Exit Sub
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    On Error Resume Next
    Set wsDest = ThisWorkbook.Worksheets("合併結果")
    On Error GoTo 0
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Worksheets.Add
        wsDest.Name = "合併結果"
    Else
        wsDest.Cells.Clear
    End If

    destRow = 1
    fileName = Dir(folderPath & "*.xlsx")
    If fileName = "" Then
        fileName = Dir(folderPath & "*.xls")
    End If

    If fileName = "" Then
        MsgBox "資料夾中找不到Excel檔案。", vbExclamation, "提示"
        Exit Sub
    End If

    isFirstFile = True

    Do While fileName <> ""
        If fileName <> ThisWorkbook.Name Then
            Set wbSource = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
            Set wsSource = wbSource.Worksheets(1)

            lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
            lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

            If isFirstFile Then
                startRow = 1
            Else
                startRow = 2
            End If

            For srcRow = startRow To lastRow
                For c = 1 To lastCol
                    wsDest.Cells(destRow, c).Value = wsSource.Cells(srcRow, c).Value
                    On Error Resume Next
                    Set cmt = wsSource.Cells(srcRow, c).Comment
                    On Error GoTo 0
                    If Not cmt Is Nothing Then
                        wsDest.Cells(destRow, c).AddComment cmt.Text
                        Set cmt = Nothing
                    End If
                Next c
                destRow = destRow + 1
            Next srcRow

            isFirstFile = False
            wbSource.Close SaveChanges:=False
        End If
        fileName = Dir()
    Loop

    wsDest.Columns.AutoFit
    MsgBox "合併完成！共 " & destRow - 1 & " 列資料（含備注）。", vbInformation, "完成"
End Sub
