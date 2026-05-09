Attribute VB_Name = "SplitSheetToCSV"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetToCSV
'功能說明: 依指定欄位的唯一值拆分工作表，每組資料另存為 CSV 格式檔案
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例資料後執行 CSV 分割
Sub TestSplitToCSV()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheetCSV(ThisWorkbook, "CSV來源資料")
    Call FillSampleCSVData(ws)
    Call SplitSheetByColumnToCSV(ws, 2)
End Sub

' 依欄位唯一值拆分工作表，各組存為 CSV 檔
' ws: 來源工作表  splitColIndex: 分割依據欄號
Sub SplitSheetByColumnToCSV(ByVal ws As Worksheet, ByVal splitColIndex As Integer)
    Dim folderPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim colCount As Long
    Dim cellValue As String
    Dim uniqueValues As Collection
    Dim alreadyAdded As Boolean
    Dim key As Variant
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim newRow As Long
    Dim copyRow As Long

    lastRow = ws.Cells(ws.Rows.Count, splitColIndex).End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "工作表無資料可分割！", vbExclamation, "警告"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 CSV 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set uniqueValues = New Collection
    For i = 2 To lastRow
        cellValue = CStr(ws.Cells(i, splitColIndex).Value)
        alreadyAdded = False
        For Each key In uniqueValues
            If key = cellValue Then
                alreadyAdded = True
                Exit For
            End If
        Next key
        If Not alreadyAdded Then uniqueValues.Add cellValue
    Next i

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each key In uniqueValues
        Set wbNew = Workbooks.Add
        Set wsNew = wbNew.Worksheets(1)
        ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Copy _
            Destination:=wsNew.Cells(1, 1)
        newRow = 2
        For copyRow = 2 To lastRow
            If CStr(ws.Cells(copyRow, splitColIndex).Value) = CStr(key) Then
                ws.Range(ws.Cells(copyRow, 1), ws.Cells(copyRow, colCount)).Copy _
                    Destination:=wsNew.Cells(newRow, 1)
                newRow = newRow + 1
            End If
        Next copyRow
        wbNew.SaveAs folderPath & CStr(key) & ".csv", xlCSV
        wbNew.Close SaveChanges:=False
    Next key

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "CSV 分割完成！共建立 " & uniqueValues.Count & " 個 CSV 檔案。", vbInformation, "完成"
End Sub

' 建立範例資料
Private Sub FillSampleCSVData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "客戶"
    ws.Range("B1").Value = "地區"
    ws.Range("C1").Value = "訂單金額"
    ws.Range("A2").Value = "客戶甲": ws.Range("B2").Value = "北區": ws.Range("C2").Value = 12000
    ws.Range("A3").Value = "客戶乙": ws.Range("B3").Value = "南區": ws.Range("C3").Value = 8500
    ws.Range("A4").Value = "客戶丙": ws.Range("B4").Value = "北區": ws.Range("C4").Value = 15000
    ws.Range("A5").Value = "客戶丁": ws.Range("B5").Value = "中區": ws.Range("C5").Value = 9300
    ws.Range("A6").Value = "客戶戊": ws.Range("B6").Value = "南區": ws.Range("C6").Value = 11200
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetCSV(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCSV = ws
End Function
