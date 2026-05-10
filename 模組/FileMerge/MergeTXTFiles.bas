'*************************************************************************************
'模組名稱: MergeTXTFiles
'功能說明: 合併指定資料夾中所有 TXT 純文字檔至一個工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************
Option Explicit

Sub MergeTXTFiles()
    Dim folderPath  As String
    Dim fileName    As String
    Dim ws          As Worksheet
    Dim nextRow     As Long
    Dim fileNo      As Integer
    Dim lineText    As String

    ' 選擇資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 TXT 檔的資料夾"
        If .Show = False Then Exit Sub
        folderPath = .SelectedItems(1) & "\\"
    End With

    ' 建立輸出工作表
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "MergedTXT"
    ws.Cells(1, 1).Value = "來源檔案"
    ws.Cells(1, 2).Value = "內容"
    nextRow = 2

    fileName = Dir(folderPath & "*.txt")
    If fileName = "" Then
        MsgBox "資料夾中找不到 TXT 檔案！", vbExclamation, "提示"
        Exit Sub
    End If

    Do While fileName <> ""
        fileNo = FreeFile
        Open folderPath & fileName For Input As #fileNo
        Do While Not EOF(fileNo)
            Line Input #fileNo, lineText
            ws.Cells(nextRow, 1).Value = fileName
            ws.Cells(nextRow, 2).Value = lineText
            nextRow = nextRow + 1
        Loop
        Close #fileNo
        fileName = Dir
    Loop

    ws.Columns("A:B").AutoFit
    MsgBox "TXT 檔案合併完成，共 " & (nextRow - 2) & " 行資料。", vbInformation, "完成"
End Sub
