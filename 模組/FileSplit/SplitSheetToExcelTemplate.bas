Attribute VB_Name = "SplitSheetToExcelTemplate"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetToExcelTemplate
'功能說明: 依指定欄位的唯一值切割工作表，並將每份資料貼入指定 .xltx 範本，
'          輸出為個別的 .xlsx 檔案至目標資料夾
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

Sub SplitByColumnToTemplate()
    Dim wsSrc       As Worksheet
    Dim templatePath As String
    Dim outputFolder As String
    Dim splitColIdx  As Long
    Dim lastRow      As Long
    Dim lastCol      As Long
    Dim dict         As Object
    Dim cell         As Range
    Dim key          As Variant
    Dim wbNew        As Workbook
    Dim wsNew        As Worksheet
    Dim destRow      As Long
    Dim r            As Long
    Dim safeKey      As String

    Set wsSrc = ActiveSheet
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "工作表沒有資料，請先確認資料範圍。", vbExclamation
        Exit Sub
    End If

    ' 輸入要分組的欄位編號
    Dim colInput As String
    colInput = InputBox("請輸入用來分割的欄位編號（例如 1 表示 A 欄）：", _
                        "設定分割欄位", "1")
    If colInput = "" Then Exit Sub
    If Not IsNumeric(colInput) Then
        MsgBox "請輸入數字欄位編號。", vbExclamation
        Exit Sub
    End If
    splitColIdx = CLng(colInput)

    ' 選擇範本檔案
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "請選擇 Excel 範本 (.xltx)"
        .Filters.Clear
        .Filters.Add "Excel 範本", "*.xltx;*.xlsx"
        If .Show <> -1 Then Exit Sub
        templatePath = .SelectedItems(1)
    End With

    ' 選擇輸出資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇輸出資料夾"
        If .Show <> -1 Then Exit Sub
        outputFolder = .SelectedItems(1)
    End With
    If Right(outputFolder, 1) <> "" Then outputFolder = outputFolder & ""

    ' 收集唯一值
    Set dict = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        key = CStr(wsSrc.Cells(r, splitColIdx).Value)
        If key <> "" And Not dict.Exists(key) Then
            dict.Add key, key
        End If
    Next r

    If dict.Count = 0 Then
        MsgBox "分割欄位沒有有效資料。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 逐一群組輸出
    For Each key In dict.Keys
        Set wbNew = Workbooks.Open(templatePath)
        Set wsNew = wbNew.Worksheets(1)

        ' 清除範本原有內容（保留格式）
        wsNew.UsedRange.ClearContents

        ' 複製標題列
        wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy
        wsNew.Range("A1").PasteSpecial Paste:=xlPasteValues

        destRow = 2
        For r = 2 To lastRow
            If CStr(wsSrc.Cells(r, splitColIdx).Value) = CStr(key) Then
                wsSrc.Range(wsSrc.Cells(r, 1), wsSrc.Cells(r, lastCol)).Copy
                wsNew.Range("A" & destRow).PasteSpecial Paste:=xlPasteValues
                destRow = destRow + 1
            End If
        Next r

        wsNew.Columns.AutoFit

        ' 以安全字元作為檔名
        safeKey = key
        Dim c As Integer
        Dim invalidChars As String
        invalidChars = "\/:*?""<>|"
        For c = 1 To Len(invalidChars)
            safeKey = Replace(safeKey, Mid(invalidChars, c, 1), "_")
        Next c

        Dim savePath As String
        savePath = outputFolder & safeKey & ".xlsx"
        Application.DisplayAlerts = False
        wbNew.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        wbNew.Close SaveChanges:=False
    Next key

    Application.ScreenUpdating = True
    Application.CutCopyMode = False

    MsgBox "切割完成！共輸出 " & dict.Count & " 個檔案至：" & outputFolder, _
           vbInformation, "完成"
End Sub
