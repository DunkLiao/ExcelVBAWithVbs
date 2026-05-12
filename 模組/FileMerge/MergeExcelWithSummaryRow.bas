Attribute VB_Name = "MergeExcelWithSummaryRow"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithSummaryRow
'功能說明: 合併資料夾內所有 Excel 檔案的第一個工作表，並在結尾附加合計列
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestMergeExcelWithSummaryRow()
    Call MergeExcelWithSummaryRow
End Sub

' 合併資料夾內 Excel 並加入合計列
Sub MergeExcelWithSummaryRow()
    On Error GoTo ErrorHandler

    Dim folderPath As String
    Dim fileName As String
    Dim wbSrc As Workbook
    Dim wsDest As Worksheet
    Dim wsSrc As Worksheet
    Dim destRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headerCopied As Boolean
    Dim numericStartRow As Long

    folderPath = GetFolderPath()
    If folderPath = "" Then
        MsgBox "未選擇資料夾，程式結束。", vbInformation, "取消"
        Exit Sub
    End If

    Set wsDest = GetOrCreateSheet(ThisWorkbook, "合併結果含合計")
    destRow = 1
    headerCopied = False
    numericStartRow = 0
    lastCol = 1

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        If fileName <> ThisWorkbook.Name Then
            Set wbSrc = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
            Set wsSrc = wbSrc.Worksheets(1)
            lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

            If Not headerCopied Then
                wsSrc.Rows(1).Copy Destination:=wsDest.Rows(destRow)
                wsDest.Rows(destRow).Font.Bold = True
                destRow = destRow + 1
                numericStartRow = destRow
                headerCopied = True
                If lastRow > 1 Then
                    wsSrc.Rows("2:" & lastRow).Copy Destination:=wsDest.Rows(destRow)
                    destRow = destRow + lastRow - 1
                End If
            Else
                If lastRow > 1 Then
                    wsSrc.Rows("2:" & lastRow).Copy Destination:=wsDest.Rows(destRow)
                    destRow = destRow + lastRow - 1
                End If
            End If

            wbSrc.Close SaveChanges:=False
        End If
        fileName = Dir()
    Loop

    ' 加入合計列
    If numericStartRow > 0 And destRow > numericStartRow Then
        Dim sumRow As Long
        sumRow = destRow
        Dim c As Long
        wsDest.Cells(sumRow, 1).Value = "合計"
        wsDest.Cells(sumRow, 1).Font.Bold = True
        For c = 2 To lastCol
            If IsNumeric(wsDest.Cells(numericStartRow, c).Value) Then
                wsDest.Cells(sumRow, c).Formula = "=SUM(" & _
                    wsDest.Cells(numericStartRow, c).Address & ":" & _
                    wsDest.Cells(sumRow - 1, c).Address & ")"
                wsDest.Cells(sumRow, c).Font.Bold = True
            End If
        Next c
    End If

    wsDest.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "合併完成，合計列已附加！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得資料夾路徑
Private Function GetFolderPath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "請選擇含有 Excel 檔案的資料夾"
    If fd.Show = -1 Then
        GetFolderPath = fd.SelectedItems(1) & Application.PathSeparator
    Else
        GetFolderPath = ""
    End If
End Function

' 取得或建立工作表
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
