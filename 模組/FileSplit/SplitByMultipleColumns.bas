Attribute VB_Name = "SplitByMultipleColumns"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByMultipleColumns
'功能說明: 依兩個欄位的組合鍵拆分工作表，每種組合存成一個 Excel 檔案
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例資料後依「地區+部門」組合分割
Sub TestSplitByMultipleColumns()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheetMulti(ThisWorkbook, "組合鍵來源")
    Call FillSampleMultiData(ws)
    Call SplitSheetByTwoColumns(ws, 1, 2)
End Sub

' 依兩個欄位的組合鍵拆分工作表
' ws: 來源工作表  col1: 第一欄號  col2: 第二欄號
Sub SplitSheetByTwoColumns(ByVal ws As Worksheet, _
                            ByVal col1 As Integer, _
                            ByVal col2 As Integer)
    Dim folderPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim colCount As Long
    Dim comboKey As String
    Dim uniqueKeys As Collection
    Dim alreadyAdded As Boolean
    Dim key As Variant
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim newRow As Long
    Dim copyRow As Long
    Dim safeName As String

    lastRow = ws.Cells(ws.Rows.Count, col1).End(xlUp).Row
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
        comboKey = CStr(ws.Cells(i, col1).Value) & "_" & CStr(ws.Cells(i, col2).Value)
        alreadyAdded = False
        For Each key In uniqueKeys
            If key = comboKey Then
                alreadyAdded = True
                Exit For
            End If
        Next key
        If Not alreadyAdded Then uniqueKeys.Add comboKey
    Next i

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each key In uniqueKeys
        Set wbNew = Workbooks.Add
        Set wsNew = wbNew.Worksheets(1)
        safeName = CStr(key)
        safeName = Replace(safeName, "/", "-")
        safeName = Replace(safeName, ":", "-")
        wsNew.Name = Left(safeName, 31)
        ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Copy _
            Destination:=wsNew.Cells(1, 1)
        newRow = 2
        For copyRow = 2 To lastRow
            comboKey = CStr(ws.Cells(copyRow, col1).Value) & "_" & CStr(ws.Cells(copyRow, col2).Value)
            If comboKey = CStr(key) Then
                ws.Range(ws.Cells(copyRow, 1), ws.Cells(copyRow, colCount)).Copy _
                    Destination:=wsNew.Cells(newRow, 1)
                newRow = newRow + 1
            End If
        Next copyRow
        wsNew.Columns.AutoFit
        wbNew.SaveAs folderPath & safeName & ".xlsx", xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
    Next key

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "組合鍵分割完成！共建立 " & uniqueKeys.Count & " 個檔案。", vbInformation, "完成"
End Sub

' 建立範例資料
Private Sub FillSampleMultiData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "姓名"
    ws.Range("D1").Value = "業績"
    ws.Range("A2").Value = "北區": ws.Range("B2").Value = "業務部": ws.Range("C2").Value = "王大明": ws.Range("D2").Value = 80000
    ws.Range("A3").Value = "北區": ws.Range("B3").Value = "財務部": ws.Range("C3").Value = "李小美": ws.Range("D3").Value = 60000
    ws.Range("A4").Value = "南區": ws.Range("B4").Value = "業務部": ws.Range("C4").Value = "張志偉": ws.Range("D4").Value = 75000
    ws.Range("A5").Value = "北區": ws.Range("B5").Value = "業務部": ws.Range("C5").Value = "陳美如": ws.Range("D5").Value = 92000
    ws.Range("A6").Value = "南區": ws.Range("B6").Value = "財務部": ws.Range("C6").Value = "林正雄": ws.Range("D6").Value = 58000
    ws.Range("A1:D1").Font.Bold = True
    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetMulti(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetMulti = ws
End Function
