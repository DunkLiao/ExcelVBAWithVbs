Option Explicit
Attribute VB_Name = "MergeExcelWithValidation"
'*************************************************************************************
'模組名稱: MergeExcelWithValidation
'功能說明: 合併多個工作簿的資料時，同時保留原始資料驗證規則
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestMergeExcelWithValidation()
    Dim selectedFiles As Variant
    selectedFiles = Application.GetOpenFilename( _
        "Excel 檔案 (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        1, "請選擇要合併的 Excel 檔案（可多選）", , True)

    If VarType(selectedFiles) = vbBoolean Then
        MsgBox "已取消操作。", vbInformation, "提示"
        Exit Sub
    End If

    Call MergeWithValidation(selectedFiles)
End Sub

Sub MergeWithValidation(ByVal fileList As Variant)
    On Error GoTo ErrorHandler

    Dim masterWb As Workbook
    Dim masterWs As Worksheet
    Set masterWb = Workbooks.Add
    Set masterWs = masterWb.Worksheets(1)
    masterWs.Name = "合併結果"

    Dim nextRow As Long
    nextRow = 1

    Dim i As Long
    For i = 1 To UBound(fileList)
        Dim srcWb As Workbook
        Set srcWb = Workbooks.Open(CStr(fileList(i)), ReadOnly:=True)

        Dim srcWs As Worksheet
        For Each srcWs In srcWb.Worksheets
            Call AppendSheetWithValidation(srcWs, masterWs, nextRow)
        Next srcWs

        srcWb.Close SaveChanges:=False
    Next i

    masterWs.Columns.AutoFit
    MsgBox "合併完成，資料驗證規則已保留！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub AppendSheetWithValidation( _
    ByVal srcWs As Worksheet, _
    ByVal destWs As Worksheet, _
    ByRef nextRow As Long)

    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    Dim srcRng As Range
    Set srcRng = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))

    Dim destCell As Range
    Set destCell = destWs.Cells(nextRow, 1)

    ' 複製資料與格式
    srcRng.Copy destCell

    ' 保留資料驗證規則
    On Error Resume Next
    srcRng.Copy
    destCell.Resize(lastRow, lastCol).PasteSpecial Paste:=xlPasteValidation
    Application.CutCopyMode = False
    On Error GoTo 0

    nextRow = nextRow + lastRow
End Sub
