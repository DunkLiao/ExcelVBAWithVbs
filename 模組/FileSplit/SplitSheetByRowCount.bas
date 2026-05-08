Option Explicit

' 依指定資料列數切割目前工作表，每個區塊另存為新活頁簿。
Public Sub SplitSheetByRowCountExample()
    On Error GoTo ErrHandler

    Dim rowsPerFile As Long
    Dim outputFolder As String

    rowsPerFile = CLng(Application.InputBox("請輸入每個檔案的資料列數", "切割列數", 100, Type:=1))
    If rowsPerFile <= 0 Then Exit Sub

    outputFolder = ThisWorkbook.Path
    If Len(outputFolder) = 0 Then
        MsgBox "請先儲存目前活頁簿，再執行切割。", vbExclamation, "提醒"
        Exit Sub
    End If

    Call SplitActiveSheetByRowCount(rowsPerFile, outputFolder)
    MsgBox "工作表切割完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "切割工作表失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub SplitActiveSheetByRowCount(ByVal rowsPerFile As Long, ByVal outputFolder As String)
    Dim srcWs As Worksheet
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim fileIndex As Long
    Dim targetRowCount As Long

    Set srcWs = ActiveSheet
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then Err.Raise vbObjectError + 510, , "目前工作表沒有可切割的資料列。"

    fileIndex = 1
    For startRow = 2 To lastRow Step rowsPerFile
        endRow = startRow + rowsPerFile - 1
        If endRow > lastRow Then endRow = lastRow

        Set newWb = Workbooks.Add(xlWBATWorksheet)
        Set newWs = newWb.Worksheets(1)
        srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(1, lastCol)).Copy Destination:=newWs.Range("A1")
        targetRowCount = endRow - startRow + 1
        srcWs.Range(srcWs.Cells(startRow, 1), srcWs.Cells(endRow, lastCol)).Copy Destination:=newWs.Range("A2")
        newWs.Columns.AutoFit
        newWb.SaveAs outputFolder & "\SplitRows_" & Format(fileIndex, "000") & ".xlsx", xlOpenXMLWorkbook
        newWb.Close SaveChanges:=False
        fileIndex = fileIndex + 1
    Next startRow
End Sub