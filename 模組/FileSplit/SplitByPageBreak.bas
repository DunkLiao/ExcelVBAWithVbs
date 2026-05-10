Attribute VB_Name = "SplitByPageBreak"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByPageBreak
'功能說明: 依據工作表水平分頁符號的位置，將資料切割並匯出為多個獨立活頁簿
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestSplitByPageBreak()
    Call SplitActiveSheetByPageBreak
End Sub

' 依分頁符號切割目前工作表並儲存為獨立活頁簿
Sub SplitActiveSheetByPageBreak()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim startRow As Long
    Dim endRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Integer
    Dim outputFolder As String
    Dim fileIndex As Integer

    Set ws = ActiveSheet

    If ws.HPageBreaks.Count = 0 Then
        MsgBox "目前工作表沒有水平分頁符號，無法切割。", vbInformation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選取輸出資料夾"
        If .Show = False Then Exit Sub
        outputFolder = .SelectedItems(1)
    End With

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    startRow = 1
    fileIndex = 1

    For i = 1 To ws.HPageBreaks.Count
        endRow = ws.HPageBreaks(i).Location.Row - 1

        If endRow >= startRow Then
            Set newWb = Workbooks.Add
            Set newWs = newWb.Worksheets(1)
            newWs.Name = ws.Name

            ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, lastCol)).Copy
            newWs.Range("A1").PasteSpecial xlPasteValues
            newWs.Range("A1").PasteSpecial xlPasteFormats
            newWs.Columns.AutoFit

            newWb.SaveAs outputFolder & "\切割_" & fileIndex & ".xlsx", xlOpenXMLWorkbook
            newWb.Close SaveChanges:=False
            fileIndex = fileIndex + 1
        End If

        startRow = ws.HPageBreaks(i).Location.Row
    Next i

    ' 處理最後一段資料
    If startRow <= lastRow Then
        Set newWb = Workbooks.Add
        Set newWs = newWb.Worksheets(1)
        newWs.Name = ws.Name

        ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol)).Copy
        newWs.Range("A1").PasteSpecial xlPasteValues
        newWs.Range("A1").PasteSpecial xlPasteFormats
        newWs.Columns.AutoFit

        newWb.SaveAs outputFolder & "\切割_" & fileIndex & ".xlsx", xlOpenXMLWorkbook
        newWb.Close SaveChanges:=False
        fileIndex = fileIndex + 1
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "切割完成！共產生 " & (fileIndex - 1) & " 個檔案。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "切割時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
