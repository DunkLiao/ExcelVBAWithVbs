Attribute VB_Name = "SplitSheetByColumn"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetByColumn
'功能說明: 將工作表依據指定欄的值分割，每個唯一值另存為獨立的Excel檔案
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口（先建立範例資料再分割）
Sub TestSplitSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "切割來源資料")
    Call FillSampleSplitData(ws)
    Call SplitSheetByColumnValue(ws, 2)
End Sub

' 依據指定欄位的唯一值分割工作表，輸出為獨立 Excel 檔案
' ws: 來源工作表  splitColIndex: 分割欄位索引
Sub SplitSheetByColumnValue(ByVal ws As Worksheet, ByVal splitColIndex As Integer)
    Dim folderPath As String
    Dim lastRow As Long
    Dim headerRow As Long
    Dim i As Long
    Dim uniqueValues As Collection
    Dim cellValue As String
    Dim alreadyAdded As Boolean
    Dim key As Variant
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim copyRow As Long
    Dim newRow As Long
    Dim colCount As Long

    headerRow = 1
    lastRow = ws.Cells(ws.Rows.Count, splitColIndex).End(xlUp).Row

    If lastRow <= headerRow Then
        MsgBox "資料列數不足，無法分割！", vbExclamation, "警告"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇分割後檔案的儲存資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    Set uniqueValues = New Collection
    For i = headerRow + 1 To lastRow
        cellValue = CStr(ws.Cells(i, splitColIndex).Value)
        alreadyAdded = False
        For Each key In uniqueValues
            If key = cellValue Then
                alreadyAdded = True
                Exit For
            End If
        Next key
        If Not alreadyAdded Then
            uniqueValues.Add cellValue
        End If
    Next i

    colCount = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each key In uniqueValues
        Set wbNew = Workbooks.Add
        Set wsNew = wbNew.Worksheets(1)
        wsNew.Name = Left(CStr(key), 31)

        ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, colCount)).Copy _
            Destination:=wsNew.Cells(1, 1)

        newRow = 2
        For copyRow = headerRow + 1 To lastRow
            If CStr(ws.Cells(copyRow, splitColIndex).Value) = CStr(key) Then
                ws.Range(ws.Cells(copyRow, 1), ws.Cells(copyRow, colCount)).Copy _
                    Destination:=wsNew.Cells(newRow, 1)
                newRow = newRow + 1
            End If
        Next copyRow

        wsNew.Columns.AutoFit
        wbNew.SaveAs folderPath & "" & CStr(key) & ".xlsx", xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
    Next key

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "分割完成！共建立 " & uniqueValues.Count & " 個檔案。", vbInformation, "完成"
End Sub

' 填入切割範例資料
Private Sub FillSampleSplitData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "薪資"
    ws.Range("A2").Value = "張三"
    ws.Range("B2").Value = "研發部"
    ws.Range("C2").Value = 60000
    ws.Range("A3").Value = "李四"
    ws.Range("B3").Value = "業務部"
    ws.Range("C3").Value = 55000
    ws.Range("A4").Value = "王五"
    ws.Range("B4").Value = "研發部"
    ws.Range("C4").Value = 65000
    ws.Range("A5").Value = "趙六"
    ws.Range("B5").Value = "行政部"
    ws.Range("C5").Value = 45000
    ws.Range("A6").Value = "孫七"
    ws.Range("B6").Value = "業務部"
    ws.Range("C6").Value = 58000
    ws.Range("A7").Value = "周八"
    ws.Range("B7").Value = "行政部"
    ws.Range("C7").Value = 47000
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表，並清除內容
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
