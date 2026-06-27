Attribute VB_Name = "MergeExcelWithBackup"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithBackup
'功能說明: 合併 Excel 檔案前自動備份原始檔案的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestMergeWithBackup()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderPath As String
    folderPath = ThisWorkbook.Path & "\MergeBackupTest"

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    ' 建立測試用範例檔案
    Dim wbSrc As Workbook
    Dim i As Integer
    For i = 1 To 3
        Set wbSrc = Workbooks.Add
        wbSrc.Worksheets(1).Range("A1").Value = "編號"
        wbSrc.Worksheets(1).Range("B1").Value = "數值"
        wbSrc.Worksheets(1).Range("A2").Value = i
        wbSrc.Worksheets(1).Range("B2").Value = i * 1000

        Dim testFile As String
        testFile = folderPath & "\SourceFile_" & i & ".xlsx"
        Application.DisplayAlerts = False
        wbSrc.SaveAs testFile, xlOpenXMLWorkbook
        wbSrc.Close False
        Application.DisplayAlerts = True
    Next i

    ' 執行合併並備份
    Call MergeExcelFilesWithBackup(folderPath, "*.xlsx")

    ' 清理測試資料夾
    On Error Resume Next
    fso.DeleteFolder folderPath, True
    On Error GoTo 0

    Set fso = Nothing
End Sub

Sub MergeExcelFilesWithBackup(ByVal folderPath As String, _
                              Optional ByVal filePattern As String = "*.xls*")
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderObj As Object
    Dim fileObj As Object

    If Not fso.FolderExists(folderPath) Then
        MsgBox "指定的資料夾不存在：" & folderPath, vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 建立備份資料夾
    Dim backupPath As String
    backupPath = folderPath & "\Backup_" & Format(Now, "yyyyMMdd_HHmmss")

    If Not fso.FolderExists(backupPath) Then
        fso.CreateFolder backupPath
    End If

    Set folderObj = fso.GetFolder(folderPath)

    ' 建立合併用的新活頁簿
    Dim wbMerge As Workbook
    Set wbMerge = Workbooks.Add
    wbMerge.Worksheets(1).Name = "合併結果"

    Dim headerRow As Integer
    headerRow = 1
    Dim firstFile As Boolean
    firstFile = True

    Dim wbSrc As Workbook
    For Each fileObj In folderObj.Files
        If Not fso.GetExtensionName(fileObj.Name) = backupPath Then
            If UCase(fileObj.Name) Like UCase(filePattern) Then
                If InStr(1, fileObj.Path, "Backup_") = 0 Then

                    ' 複製原始檔案到備份目錄
                    fso.CopyFile fileObj.Path, backupPath & "" & fileObj.Name, True

                    ' 開啟並讀取
                    Set wbSrc = Workbooks.Open(fileObj.Path, ReadOnly:=True)

                    Dim lastRow As Long
                    lastRow = wbSrc.Worksheets(1).Cells( _
                        wbSrc.Worksheets(1).Rows.Count, 1).End(xlUp).Row

                    If firstFile Then
                        wbSrc.Worksheets(1).Range("A1").CurrentRegion.Copy _
                            wbMerge.Worksheets("合併結果").Range("A1")
                        headerRow = wbMerge.Worksheets("合併結果").Cells( _
                            wbMerge.Worksheets("合併結果").Rows.Count, 1).End(xlUp).Row + 1
                        firstFile = False
                    Else
                        wbSrc.Worksheets(1).Range("A2:A" & lastRow).EntireRow.Copy _
                            wbMerge.Worksheets("合併結果").Cells(headerRow, 1)

                        ' 加入來源標記
                        wbMerge.Worksheets("合併結果").Cells(headerRow, 1).Value = _
                            fso.GetBaseName(fileObj.Name) & "_" & _
                            wbMerge.Worksheets("合併結果").Cells(headerRow, 1).Value
                        headerRow = headerRow + lastRow - 1
                    End If

                    wbSrc.Close False

                End If
            End If
        End If
    Next fileObj

    wbMerge.Worksheets("合併結果").Columns.AutoFit

    MsgBox "合併完成！備份已存放至：" & vbCrLf & backupPath, vbInformation, "完成"

    Set wbSrc = Nothing
    Set wbMerge = Nothing
    Set fileObj = Nothing
    Set folderObj = Nothing
    Set fso = Nothing
End Sub
