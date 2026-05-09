Attribute VB_Name = "SplitByOddEvenRow"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByOddEvenRow
'功能說明: 將工作表依資料列的奇偶性拆分為兩個 Excel 檔案
'          奇數資料列與偶數資料列各存一個檔案，均保留標題列
'          適用於交錯式資料或輪替分組場景
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例資料後執行奇偶分割
Sub TestSplitByOddEven()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheetOddEven(ThisWorkbook, "奇偶列來源")
    Call FillSampleOddEvenData(ws)
    Call SplitSheetByOddEvenRow(ws)
End Sub

' 將工作表資料依奇偶行拆分為兩個 Excel 檔案
' ws: 來源工作表（第 1 列為標題，資料從第 2 列開始）
Sub SplitSheetByOddEvenRow(ByVal ws As Worksheet)
    Dim folderPath As String
    Dim lastRow As Long
    Dim colCount As Long
    Dim wbOdd As Workbook
    Dim wbEven As Workbook
    Dim wsOdd As Worksheet
    Dim wsEven As Worksheet
    Dim rowOdd As Long
    Dim rowEven As Long
    Dim i As Long
    Dim dataIndex As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
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

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbOdd = Workbooks.Add
    Set wsOdd = wbOdd.Worksheets(1)
    wsOdd.Name = "奇數列"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Copy Destination:=wsOdd.Cells(1, 1)

    Set wbEven = Workbooks.Add
    Set wsEven = wbEven.Worksheets(1)
    wsEven.Name = "偶數列"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Copy Destination:=wsEven.Cells(1, 1)

    rowOdd = 2
    rowEven = 2
    dataIndex = 0

    For i = 2 To lastRow
        dataIndex = dataIndex + 1
        If dataIndex Mod 2 = 1 Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, colCount)).Copy _
                Destination:=wsOdd.Cells(rowOdd, 1)
            rowOdd = rowOdd + 1
        Else
            ws.Range(ws.Cells(i, 1), ws.Cells(i, colCount)).Copy _
                Destination:=wsEven.Cells(rowEven, 1)
            rowEven = rowEven + 1
        End If
    Next i

    wsOdd.Columns.AutoFit
    wbOdd.SaveAs folderPath & "OddRows.xlsx", xlOpenXMLWorkbook
    wbOdd.Close SaveChanges:=False

    wsEven.Columns.AutoFit
    wbEven.SaveAs folderPath & "EvenRows.xlsx", xlOpenXMLWorkbook
    wbEven.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "奇偶列分割完成！已輸出 OddRows.xlsx 與 EvenRows.xlsx。", vbInformation, "完成"
End Sub

' 建立範例資料
Private Sub FillSampleOddEvenData(ByVal ws As Worksheet)
    Dim i As Integer
    Dim names(1 To 8) As String
    ws.Cells.Clear
    ws.Range("A1").Value = "編號"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "分數"
    names(1) = "甲": names(2) = "乙": names(3) = "丙": names(4) = "丁"
    names(5) = "戊": names(6) = "己": names(7) = "庚": names(8) = "辛"
    For i = 1 To 8
        ws.Cells(i + 1, 1).Value = i
        ws.Cells(i + 1, 2).Value = names(i)
        ws.Cells(i + 1, 3).Value = 60 + i * 4
    Next i
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetOddEven(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetOddEven = ws
End Function
