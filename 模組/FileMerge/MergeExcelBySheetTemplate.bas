Attribute VB_Name = "MergeExcelBySheetTemplate"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelBySheetTemplate
'功能說明: 依據範本工作表的欄位結構，合併多個 Excel 活頁簿中相同工作表名稱的資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點：選取資料夾並依範本合併
Sub TestMergeExcelBySheetTemplate()
    Dim folderPath As String
    Dim templateSheet As String

    templateSheet = InputBox("請輸入要合併的工作表名稱：", "指定工作表", "Sheet1")
    If templateSheet = "" Then Exit Sub

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選取包含要合併 Excel 檔案的資料夾"
        If .Show = False Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    Call MergeExcelBySheetTemplate(folderPath, templateSheet)
End Sub

' 依範本工作表名稱合併指定資料夾中所有 Excel 檔案的資料
' folderPath:    來源資料夾路徑
' templateSheet: 要合併的工作表名稱
Sub MergeExcelBySheetTemplate(ByVal folderPath As String, ByVal templateSheet As String)
    On Error GoTo ErrorHandler

    Dim outputWb As Workbook
    Dim outputWs As Worksheet
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim fileName As String
    Dim outputRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim fileCount As Integer

    Set outputWb = Workbooks.Add
    Set outputWs = outputWb.Worksheets(1)
    outputWs.Name = "合併結果"
    outputRow = 1
    fileCount = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    fileName = Dir(folderPath & "\*.xlsx")
    Do While fileName <> ""
        If outputWb.FullName <> folderPath & "" & fileName Then
            Set srcWb = Workbooks.Open(folderPath & "" & fileName, ReadOnly:=True)
            On Error Resume Next
            Set srcWs = srcWb.Worksheets(templateSheet)
            On Error GoTo 0

            If Not srcWs Is Nothing Then
                lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
                lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

                If lastRow >= 1 And lastCol >= 1 Then
                    ' 第一個檔案複製標頭列
                    If fileCount = 0 And outputRow = 1 Then
                        srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(1, lastCol)).Copy
                        outputWs.Cells(outputRow, 1).PasteSpecial xlPasteValues
                        outputRow = outputRow + 1
                    End If

                    ' 複製資料列（跳過標頭）
                    If lastRow >= 2 Then
                        srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(lastRow, lastCol)).Copy
                        outputWs.Cells(outputRow, 1).PasteSpecial xlPasteValues
                        outputRow = outputRow + (lastRow - 1)
                    End If

                    fileCount = fileCount + 1
                End If
                Set srcWs = Nothing
            End If

            srcWb.Close SaveChanges:=False
        End If
        fileName = Dir
    Loop

    outputWs.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "合併完成！共處理 " & fileCount & " 個檔案，" & _
           "合計 " & (outputRow - 2) & " 筆資料。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
