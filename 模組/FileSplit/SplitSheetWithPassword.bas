Attribute VB_Name = "SplitSheetWithPassword"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetWithPassword
'功能說明: 依欄位唯一值拆分工作表，並對每個輸出的 Excel 檔案設定開啟密碼保護
'          適合需要分發給不同收件人且要求保密的情境
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例資料後執行加密分割
Sub TestSplitWithPassword()
    Dim ws As Worksheet
    Dim pwd As String
    Set ws = GetOrCreateSheetPwd(ThisWorkbook, "加密分割來源")
    Call FillSamplePwdData(ws)
    pwd = InputBox("請輸入每個分割檔案的開啟密碼：", "設定密碼", "P@ssw0rd")
    If Len(Trim(pwd)) = 0 Then
        MsgBox "密碼不可為空白，操作取消。", vbExclamation, "取消"
        Exit Sub
    End If
    Call SplitSheetByColumnWithPassword(ws, 2, pwd)
End Sub

' 依欄位唯一值拆分工作表並以密碼保護輸出檔
' ws: 來源工作表  splitColIndex: 分割欄號  filePassword: 檔案開啟密碼
Sub SplitSheetByColumnWithPassword(ByVal ws As Worksheet, _
                                    ByVal splitColIndex As Integer, _
                                    ByVal filePassword As String)
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
        .Title = "請選擇加密檔案輸出資料夾"
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
        wsNew.Name = Left(CStr(key), 31)
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
        wsNew.Columns.AutoFit
        ' 以密碼保護儲存 Excel 檔案
        wbNew.SaveAs folderPath & CStr(key) & ".xlsx", _
            FileFormat:=xlOpenXMLWorkbook, _
            Password:=filePassword
        wbNew.Close SaveChanges:=False
    Next key

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "加密分割完成！共建立 " & uniqueValues.Count & " 個加密檔案。", vbInformation, "完成"
End Sub

' 建立範例資料
Private Sub FillSamplePwdData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "員工編號"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "薪資"
    ws.Range("D1").Value = "考績"
    ws.Range("A2").Value = "E001": ws.Range("B2").Value = "財務部": ws.Range("C2").Value = 55000: ws.Range("D2").Value = "A"
    ws.Range("A3").Value = "E002": ws.Range("B3").Value = "研發部": ws.Range("C3").Value = 72000: ws.Range("D3").Value = "B"
    ws.Range("A4").Value = "E003": ws.Range("B4").Value = "財務部": ws.Range("C4").Value = 48000: ws.Range("D4").Value = "A"
    ws.Range("A5").Value = "E004": ws.Range("B5").Value = "業務部": ws.Range("C5").Value = 61000: ws.Range("D5").Value = "C"
    ws.Range("A6").Value = "E005": ws.Range("B6").Value = "研發部": ws.Range("C6").Value = 83000: ws.Range("D6").Value = "A"
    ws.Range("A1:D1").Font.Bold = True
    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetPwd(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetPwd = ws
End Function
