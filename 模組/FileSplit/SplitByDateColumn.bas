Attribute VB_Name = "SplitByDateColumn"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByDateColumn
'功能說明: 依指定日期欄位的年月，將工作表拆分為多個 Excel 檔案
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例資料後執行拆分
Sub TestSplitByDate()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheetByDate(ThisWorkbook, "日期來源資料")
    Call FillSampleDateData(ws)
    Call SplitSheetByDateColumn(ws, 1)
End Sub

' 依日期欄（年月）拆分工作表，每個年月存成獨立 Excel 檔案
' ws: 來源工作表  dateColIndex: 日期所在欄號
Sub SplitSheetByDateColumn(ByVal ws As Worksheet, ByVal dateColIndex As Integer)
    Dim folderPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim colCount As Long
    Dim yearMonth As String
    Dim uniqueKeys As Collection
    Dim alreadyAdded As Boolean
    Dim key As Variant
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim newRow As Long
    Dim copyRow As Long

    lastRow = ws.Cells(ws.Rows.Count, dateColIndex).End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "工作表無資料可分割！", vbExclamation, "警告"
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

    Set uniqueKeys = New Collection
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, dateColIndex).Value) Then
            yearMonth = Format(CDate(ws.Cells(i, dateColIndex).Value), "YYYY-MM")
            alreadyAdded = False
            For Each key In uniqueKeys
                If key = yearMonth Then
                    alreadyAdded = True
                    Exit For
                End If
            Next key
            If Not alreadyAdded Then uniqueKeys.Add yearMonth
        End If
    Next i

    If uniqueKeys.Count = 0 Then
        MsgBox "找不到有效的日期資料！", vbExclamation, "警告"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each key In uniqueKeys
        Set wbNew = Workbooks.Add
        Set wsNew = wbNew.Worksheets(1)
        wsNew.Name = Left(CStr(key), 31)
        ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Copy _
            Destination:=wsNew.Cells(1, 1)
        newRow = 2
        For copyRow = 2 To lastRow
            If IsDate(ws.Cells(copyRow, dateColIndex).Value) Then
                If Format(CDate(ws.Cells(copyRow, dateColIndex).Value), "YYYY-MM") = CStr(key) Then
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
    MsgBox "拆分完成！共建立 " & uniqueKeys.Count & " 個年月檔案。", vbInformation, "完成"
End Sub

' 建立範例日期資料
Private Sub FillSampleDateData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "品名"
    ws.Range("C1").Value = "金額"
    ws.Range("A2").Value = CDate("2026/01/05")
    ws.Range("B2").Value = "商品A"
    ws.Range("C2").Value = 3000
    ws.Range("A3").Value = CDate("2026/01/12")
    ws.Range("B3").Value = "商品B"
    ws.Range("C3").Value = 1500
    ws.Range("A4").Value = CDate("2026/02/03")
    ws.Range("B4").Value = "商品C"
    ws.Range("C4").Value = 4200
    ws.Range("A5").Value = CDate("2026/02/18")
    ws.Range("B5").Value = "商品D"
    ws.Range("C5").Value = 2100
    ws.Range("A6").Value = CDate("2026/03/07")
    ws.Range("B6").Value = "商品E"
    ws.Range("C6").Value = 5500
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetByDate(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetByDate = ws
End Function
