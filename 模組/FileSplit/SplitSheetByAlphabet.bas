Attribute VB_Name = "SplitSheetByAlphabet"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetByAlphabet
'功能說明: 依指定欄位的首字母（A-Z）分組，將工作表拆分為多個 Excel 檔案
'          非英文字母的值統一歸入 Others 群組
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例資料後依首字母分割
Sub TestSplitByAlphabet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheetAlpha(ThisWorkbook, "字母分組來源")
    Call FillSampleAlphaData(ws)
    Call SplitSheetByFirstLetter(ws, 1)
End Sub

' 依指定欄位的首字母分割工作表
' ws: 來源工作表  keyColIndex: 鍵值欄號
Sub SplitSheetByFirstLetter(ByVal ws As Worksheet, ByVal keyColIndex As Integer)
    Dim folderPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim colCount As Long
    Dim firstLetter As String
    Dim uniqueLetters As Collection
    Dim alreadyAdded As Boolean
    Dim key As Variant
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim newRow As Long
    Dim copyRow As Long
    Dim cellVal As String

    lastRow = ws.Cells(ws.Rows.Count, keyColIndex).End(xlUp).Row
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

    Set uniqueLetters = New Collection
    For i = 2 To lastRow
        cellVal = CStr(ws.Cells(i, keyColIndex).Value)
        If Len(cellVal) > 0 Then
            firstLetter = UCase(Left(cellVal, 1))
            If firstLetter < "A" Or firstLetter > "Z" Then firstLetter = "Others"
            alreadyAdded = False
            For Each key In uniqueLetters
                If key = firstLetter Then
                    alreadyAdded = True
                    Exit For
                End If
            Next key
            If Not alreadyAdded Then uniqueLetters.Add firstLetter
        End If
    Next i

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each key In uniqueLetters
        Set wbNew = Workbooks.Add
        Set wsNew = wbNew.Worksheets(1)
        wsNew.Name = Left(CStr(key), 31)
        ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Copy _
            Destination:=wsNew.Cells(1, 1)
        newRow = 2
        For copyRow = 2 To lastRow
            cellVal = CStr(ws.Cells(copyRow, keyColIndex).Value)
            If Len(cellVal) > 0 Then
                firstLetter = UCase(Left(cellVal, 1))
                If firstLetter < "A" Or firstLetter > "Z" Then firstLetter = "Others"
                If firstLetter = CStr(key) Then
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
    MsgBox "首字母分割完成！共建立 " & uniqueLetters.Count & " 個檔案。", vbInformation, "完成"
End Sub

' 建立範例資料
Private Sub FillSampleAlphaData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "公司名稱"
    ws.Range("B1").Value = "聯絡人"
    ws.Range("C1").Value = "電話"
    ws.Range("A2").Value = "Apple Corp": ws.Range("B2").Value = "John": ws.Range("C2").Value = "02-1234-5678"
    ws.Range("A3").Value = "Beta Solutions": ws.Range("B3").Value = "Mary": ws.Range("C3").Value = "02-2345-6789"
    ws.Range("A4").Value = "Ace Trading": ws.Range("B4").Value = "David": ws.Range("C4").Value = "02-3456-7890"
    ws.Range("A5").Value = "Cherry Design": ws.Range("B5").Value = "Susan": ws.Range("C5").Value = "02-4567-8901"
    ws.Range("A6").Value = "Big Data Inc": ws.Range("B6").Value = "Kevin": ws.Range("C6").Value = "02-5678-9012"
    ws.Range("A7").Value = "協大工業": ws.Range("B7").Value = "陳大偉": ws.Range("C7").Value = "04-6789-0123"
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetAlpha(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetAlpha = ws
End Function
