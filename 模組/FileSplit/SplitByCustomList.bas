Attribute VB_Name = "SplitByCustomList"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByCustomList
'功能說明: 依另一工作表定義的自訂清單分組，將來源工作表拆分為多個 Excel 檔案
'          清單工作表：A 欄為群組名稱，B 欄起為該群組包含的欄位值（可多個）
'          不符合任何群組的列存入 Others.xlsx
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例資料與自訂清單後執行分割
Sub TestSplitByCustomList()
    Dim wsSrc As Worksheet
    Dim wsMap As Worksheet
    Set wsSrc = GetOrCreateSheetCustom(ThisWorkbook, "自訂清單來源")
    Set wsMap = GetOrCreateSheetCustom(ThisWorkbook, "分組清單定義")
    Call FillSampleCustomData(wsSrc)
    Call FillSampleCustomList(wsMap)
    Call SplitSheetByCustomGroupList(wsSrc, wsMap, 2)
End Sub

' 依自訂群組清單拆分工作表
' wsSrc: 來源工作表  wsMap: 群組定義工作表  splitColIndex: 比對欄號
Sub SplitSheetByCustomGroupList(ByVal wsSrc As Worksheet, _
                                 ByVal wsMap As Worksheet, _
                                 ByVal splitColIndex As Integer)
    Dim folderPath As String
    Dim lastRowSrc As Long
    Dim lastRowMap As Long
    Dim lastColMap As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colCount As Long
    Dim groupName As String
    Dim groupValue As String
    Dim cellValue As String
    Dim groupFound As String
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim newRow As Long
    Dim groupCount As Long

    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, splitColIndex).End(xlUp).Row
    If lastRowSrc <= 1 Then
        MsgBox "來源工作表無資料！", vbExclamation, "警告"
        Exit Sub
    End If

    lastRowMap = wsMap.Cells(wsMap.Rows.Count, 1).End(xlUp).Row
    If lastRowMap < 1 Then
        MsgBox "群組定義工作表無資料！", vbExclamation, "警告"
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

    colCount = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    groupCount = 0

    ' 逐一處理每個群組
    For i = 1 To lastRowMap
        groupName = CStr(wsMap.Cells(i, 1).Value)
        If Len(Trim(groupName)) = 0 Then GoTo NextGroup

        Set wbNew = Workbooks.Add
        Set wsNew = wbNew.Worksheets(1)
        wsNew.Name = Left(groupName, 31)
        wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, colCount)).Copy _
            Destination:=wsNew.Cells(1, 1)
        newRow = 2
        lastColMap = wsMap.Cells(i, wsMap.Columns.Count).End(xlToLeft).Column

        For k = 2 To lastRowSrc
            cellValue = CStr(wsSrc.Cells(k, splitColIndex).Value)
            For j = 2 To lastColMap
                groupValue = CStr(wsMap.Cells(i, j).Value)
                If cellValue = groupValue Then
                    wsSrc.Range(wsSrc.Cells(k, 1), wsSrc.Cells(k, colCount)).Copy _
                        Destination:=wsNew.Cells(newRow, 1)
                    newRow = newRow + 1
                    Exit For
                End If
            Next j
        Next k

        wsNew.Columns.AutoFit
        wbNew.SaveAs folderPath & groupName & ".xlsx", xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
        groupCount = groupCount + 1
NextGroup:
    Next i

    ' 找出不屬於任何群組的列，存入 Others
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Worksheets(1)
    wsNew.Name = "Others"
    wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, colCount)).Copy _
        Destination:=wsNew.Cells(1, 1)
    newRow = 2

    For k = 2 To lastRowSrc
        cellValue = CStr(wsSrc.Cells(k, splitColIndex).Value)
        groupFound = ""
        For i = 1 To lastRowMap
            lastColMap = wsMap.Cells(i, wsMap.Columns.Count).End(xlToLeft).Column
            For j = 2 To lastColMap
                If cellValue = CStr(wsMap.Cells(i, j).Value) Then
                    groupFound = CStr(wsMap.Cells(i, 1).Value)
                    Exit For
                End If
            Next j
            If Len(groupFound) > 0 Then Exit For
        Next i
        If Len(groupFound) = 0 Then
            wsSrc.Range(wsSrc.Cells(k, 1), wsSrc.Cells(k, colCount)).Copy _
                Destination:=wsNew.Cells(newRow, 1)
            newRow = newRow + 1
        End If
    Next k

    If newRow > 2 Then
        wsNew.Columns.AutoFit
        wbNew.SaveAs folderPath & "Others.xlsx", xlOpenXMLWorkbook
        groupCount = groupCount + 1
    End If
    wbNew.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "自訂清單分割完成！共建立 " & groupCount & " 個檔案。", vbInformation, "完成"
End Sub

' 建立來源範例資料
Private Sub FillSampleCustomData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "商品"
    ws.Range("B1").Value = "類別"
    ws.Range("C1").Value = "售價"
    ws.Range("A2").Value = "蘋果": ws.Range("B2").Value = "水果": ws.Range("C2").Value = 50
    ws.Range("A3").Value = "筆記本": ws.Range("B3").Value = "文具": ws.Range("C3").Value = 35
    ws.Range("A4").Value = "香蕉": ws.Range("B4").Value = "水果": ws.Range("C4").Value = 30
    ws.Range("A5").Value = "原子筆": ws.Range("B5").Value = "文具": ws.Range("C5").Value = 15
    ws.Range("A6").Value = "手機": ws.Range("B6").Value = "電子": ws.Range("C6").Value = 15000
    ws.Range("A7").Value = "橘子": ws.Range("B7").Value = "水果": ws.Range("C7").Value = 40
    ws.Range("A8").Value = "平板": ws.Range("B8").Value = "電子": ws.Range("C8").Value = 20000
    ws.Range("A9").Value = "雨傘": ws.Range("B9").Value = "雜貨": ws.Range("C9").Value = 280
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' 建立自訂群組清單定義（A欄=群組名，B欄起=所含類別值）
Private Sub FillSampleCustomList(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "鮮果類"
    ws.Range("B1").Value = "水果"
    ws.Range("A2").Value = "辦公用品"
    ws.Range("B2").Value = "文具"
    ws.Range("A3").Value = "3C商品"
    ws.Range("B3").Value = "電子"
    ws.Columns("A:B").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetCustom(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCustom = ws
End Function
