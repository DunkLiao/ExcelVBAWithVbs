Attribute VB_Name = "SplitByNumericRange"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByNumericRange
'功能說明: 依指定數值欄的區間（每段 rangeSize 為一組）拆分工作表，
'          每個區間存成一個 Excel 檔案
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例資料後依薪資區間（每段 10000）分割
Sub TestSplitByNumericRange()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheetNumeric(ThisWorkbook, "數值區間來源")
    Call FillSampleNumericData(ws)
    Call SplitSheetByNumericRange(ws, 3, 10000)
End Sub

' 依數值欄位區間拆分工作表
' ws: 來源工作表  numColIndex: 數值欄號  rangeSize: 每段區間大小
Sub SplitSheetByNumericRange(ByVal ws As Worksheet, _
                              ByVal numColIndex As Integer, _
                              ByVal rangeSize As Double)
    Dim folderPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim colCount As Long
    Dim cellNum As Double
    Dim bucketKey As String
    Dim bucketKey2 As String
    Dim bucketLow As Long
    Dim bucketHigh As Long
    Dim uniqueBuckets As Collection
    Dim alreadyAdded As Boolean
    Dim key As Variant
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim newRow As Long
    Dim copyRow As Long
    Dim cellNum2 As Double

    lastRow = ws.Cells(ws.Rows.Count, numColIndex).End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "工作表無資料可分割！", vbExclamation, "警告"
        Exit Sub
    End If

    If rangeSize <= 0 Then
        MsgBox "區間大小必須大於 0！", vbExclamation, "警告"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set uniqueBuckets = New Collection
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, numColIndex).Value) Then
            cellNum = CDbl(ws.Cells(i, numColIndex).Value)
            bucketLow = Int(cellNum / rangeSize) * CLng(rangeSize)
            bucketHigh = bucketLow + CLng(rangeSize) - 1
            bucketKey = CStr(bucketLow) & "-" & CStr(bucketHigh)
            alreadyAdded = False
            For Each key In uniqueBuckets
                If key = bucketKey Then
                    alreadyAdded = True
                    Exit For
                End If
            Next key
            If Not alreadyAdded Then uniqueBuckets.Add bucketKey
        End If
    Next i

    If uniqueBuckets.Count = 0 Then
        MsgBox "找不到有效的數值資料！", vbExclamation, "警告"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each key In uniqueBuckets
        Set wbNew = Workbooks.Add
        Set wsNew = wbNew.Worksheets(1)
        wsNew.Name = Left(CStr(key), 31)
        ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Copy _
            Destination:=wsNew.Cells(1, 1)
        newRow = 2
        For copyRow = 2 To lastRow
            If IsNumeric(ws.Cells(copyRow, numColIndex).Value) Then
                cellNum2 = CDbl(ws.Cells(copyRow, numColIndex).Value)
                bucketLow = Int(cellNum2 / rangeSize) * CLng(rangeSize)
                bucketHigh = bucketLow + CLng(rangeSize) - 1
                bucketKey2 = CStr(bucketLow) & "-" & CStr(bucketHigh)
                If bucketKey2 = CStr(key) Then
                    ws.Range(ws.Cells(copyRow, 1), ws.Cells(copyRow, colCount)).Copy _
                        Destination:=wsNew.Cells(newRow, 1)
                    newRow = newRow + 1
                End If
            End If
        Next copyRow
        wsNew.Columns.AutoFit
        wbNew.SaveAs folderPath & CStr(key) & ".xlsx", xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
    Next key

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "數值區間分割完成！共建立 " & uniqueBuckets.Count & " 個檔案。", vbInformation, "完成"
End Sub

' 建立範例資料
Private Sub FillSampleNumericData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "員工"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "薪資"
    ws.Range("A2").Value = "甲": ws.Range("B2").Value = "研發": ws.Range("C2").Value = 35000
    ws.Range("A3").Value = "乙": ws.Range("B3").Value = "業務": ws.Range("C3").Value = 28000
    ws.Range("A4").Value = "丙": ws.Range("B4").Value = "研發": ws.Range("C4").Value = 52000
    ws.Range("A5").Value = "丁": ws.Range("B5").Value = "行政": ws.Range("C5").Value = 42000
    ws.Range("A6").Value = "戊": ws.Range("B6").Value = "業務": ws.Range("C6").Value = 31000
    ws.Range("A7").Value = "己": ws.Range("B7").Value = "研發": ws.Range("C7").Value = 65000
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetNumeric(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetNumeric = ws
End Function
