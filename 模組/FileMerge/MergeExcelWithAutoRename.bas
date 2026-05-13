Attribute VB_Name = "MergeExcelWithAutoRename"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithAutoRename
'功能說明: 合併指定資料夾中所有 Excel 檔案至主活頁簿，
'          並在工作表名稱重複時自動加上序號重新命名
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub MergeExcelWithAutoRename()
    Dim folderPath  As String
    Dim fileName    As String
    Dim srcWb       As Workbook
    Dim srcWs       As Worksheet
    Dim dstWb       As Workbook
    Dim newName     As String
    Dim counter     As Integer

    folderPath = Application.GetOpenFilename( _
        FileFilter:="Excel 檔案 (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm", _
        Title:="選取任一目標資料夾內的檔案以確認資料夾位置")

    If folderPath = "False" Then
        MsgBox "已取消操作。", vbExclamation, "取消"
        Exit Sub
    End If

    folderPath = Left(folderPath, InStrRev(folderPath, "\"))
    Set dstWb = ThisWorkbook

    fileName = Dir(folderPath & "*.xls*")
    If fileName = "" Then
        MsgBox "資料夾內沒有找到 Excel 檔案。", vbExclamation, "提示"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Do While fileName <> ""
        Dim fullPath As String
        fullPath = folderPath & fileName

        If fullPath <> dstWb.FullName Then
            Set srcWb = Workbooks.Open(Filename:=fullPath, ReadOnly:=True)

            For Each srcWs In srcWb.Worksheets
                newName = srcWs.Name
                counter = 1

                Do While SheetExists(dstWb, newName)
                    newName = srcWs.Name & "_" & counter
                    counter = counter + 1
                Loop

                srcWs.Copy After:=dstWb.Sheets(dstWb.Sheets.Count)
                dstWb.Sheets(dstWb.Sheets.Count).Name = newName
            Next srcWs

            srcWb.Close SaveChanges:=False
        End If

        fileName = Dir
    Loop

    Application.ScreenUpdating = True
    MsgBox "合併完成，所有重複工作表名稱已自動加上序號。", vbInformation, "完成"
End Sub

' 判斷工作表是否存在
Private Function SheetExists(ByVal wb As Workbook, ByVal sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function
