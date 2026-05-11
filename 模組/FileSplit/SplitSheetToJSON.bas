Attribute VB_Name = "SplitSheetToJSON"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetToJSON
'功能說明: 將每個工作表的資料分別轉存為獨立的 JSON 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestSplitSheetToJSON()
    Call SplitSheetsToJSON
End Sub

' 將每個工作表轉存為獨立 JSON 檔案
Sub SplitSheetsToJSON()
    On Error GoTo ErrorHandler

    Dim outputFolder As String
    Dim ws As Worksheet
    Dim fileStream As Object
    Dim jsonContent As String
    Dim savedCount As Integer

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 JSON 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        outputFolder = .SelectedItems(1)
    End With

    savedCount = 0
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        Dim lastRow As Long
        Dim lastCol As Long
        Dim r As Long
        Dim c As Long
        Dim headers() As String
        Dim cellVal As String

        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        If lastRow < 2 Or lastCol < 1 Then GoTo NextSheet

        ReDim headers(1 To lastCol)
        For c = 1 To lastCol
            headers(c) = CStr(ws.Cells(1, c).Value)
        Next c

        jsonContent = "[" & vbCrLf
        For r = 2 To lastRow
            jsonContent = jsonContent & "  {"
            For c = 1 To lastCol
                cellVal = CStr(ws.Cells(r, c).Value)
                cellVal = Replace(cellVal, "\", "\\")
                cellVal = Replace(cellVal, Chr(34), "\" & Chr(34))
                jsonContent = jsonContent & Chr(34) & headers(c) & Chr(34) & ":" & _
                              Chr(34) & cellVal & Chr(34)
                If c < lastCol Then jsonContent = jsonContent & ","
            Next c
            jsonContent = jsonContent & "}"
            If r < lastRow Then jsonContent = jsonContent & ","
            jsonContent = jsonContent & vbCrLf
        Next r
        jsonContent = jsonContent & "]"

        Dim outputPath As String
        outputPath = outputFolder & "\" & ws.Name & ".json"

        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Open
        fileStream.Type = 2
        fileStream.Charset = "UTF-8"
        fileStream.WriteText jsonContent
        fileStream.SaveToFile outputPath, 2
        fileStream.Close

        savedCount = savedCount + 1

NextSheet:
    Next ws

    Application.ScreenUpdating = True
    MsgBox "已將 " & savedCount & " 個工作表轉存為 JSON 檔案至：" & vbCrLf & outputFolder, _
           vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "轉存 JSON 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
