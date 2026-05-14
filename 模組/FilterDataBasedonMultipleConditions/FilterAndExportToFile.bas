Attribute VB_Name = "FilterAndExportToFile"
Option Explicit
'*************************************************************************************
'模組名稱: FilterAndExportToFile
'功能說明: 依多重條件篹選資料後，將結果匯出為 CSV 檔案的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點：篹選部門=業務部 且 薪資>=40000，匯出 CSV
Sub TestFilterAndExportToFile()
    Dim ws      As Worksheet
    Dim outPath As String

    Set ws = GetOrCreateFilterExportSheet("篹選匯出範例")
    Call FillFilterSampleData(ws)

    outPath = Environ("USERPROFILE") & "\Desktop\FilterExport_" & _
              Format(Now, "yyyymmdd_HHmmss") & ".csv"

    Call FilterAndExportCSV(ws, outPath, "業務部", 40000)
End Sub

' 依部門與最低薪資條件篹選，並將結果匯出為 CSV
' ws: 來源工作表
' csvPath: 輸出 CSV 路徑
' deptFilter: 部門篹選值（空字串表示不篹選）
' minSalary: 最低薪資（0 表示不限）
Sub FilterAndExportCSV(ByVal ws As Worksheet, ByVal csvPath As String, _
                        ByVal deptFilter As String, ByVal minSalary As Long)
    On Error GoTo ErrorHandler

    Dim fso       As Object
    Dim ts        As Object
    Dim lastRow   As Long
    Dim lastCol   As Long
    Dim r         As Long
    Dim c         As Long
    Dim dept      As String
    Dim salary    As Long
    Dim lineStr   As String
    Dim count     As Long
    Dim deptCol   As Long
    Dim salaryCol As Long
    Dim colName   As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    deptCol = 0
    salaryCol = 0
    For c = 1 To lastCol
        colName = CStr(ws.Cells(1, c).Value)
        If colName = "部門" Then deptCol = c
        If colName = "薪資" Then salaryCol = c
    Next c

    Set ts = fso.CreateTextFile(csvPath, True, False)

    lineStr = ""
    For c = 1 To lastCol
        If c > 1 Then lineStr = lineStr & ","
        lineStr = lineStr & EscapeCSVField(CStr(ws.Cells(1, c).Value))
    Next c
    ts.WriteLine lineStr

    count = 0
    For r = 2 To lastRow
        If deptCol > 0 And deptFilter <> "" Then
            dept = CStr(ws.Cells(r, deptCol).Value)
            If dept <> deptFilter Then GoTo SkipRow
        End If

        If salaryCol > 0 And minSalary > 0 Then
            If IsNumeric(ws.Cells(r, salaryCol).Value) Then
                salary = CLng(ws.Cells(r, salaryCol).Value)
                If salary < minSalary Then GoTo SkipRow
            Else
                GoTo SkipRow
            End If
        End If

        lineStr = ""
        For c = 1 To lastCol
            If c > 1 Then lineStr = lineStr & ","
            lineStr = lineStr & EscapeCSVField(CStr(ws.Cells(r, c).Value))
        Next c
        ts.WriteLine lineStr
        count = count + 1

SkipRow:
    Next r

    ts.Close
    MsgBox "篹選匯出完成！共 " & count & " 筆資料。" & Chr(10) & _
           "檔案路徑：" & csvPath, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    If Not ts Is Nothing Then
        On Error Resume Next
        ts.Close
        On Error GoTo 0
    End If
    MsgBox "篹選匯出時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' CSV 欄位逸出（含逗號或雙引號時加引號）
Private Function EscapeCSVField(ByVal val As String) As String
    If InStr(val, ",") > 0 Or InStr(val, Chr(34)) > 0 Or InStr(val, Chr(10)) > 0 Then
        val = Replace(val, Chr(34), Chr(34) & Chr(34))
        EscapeCSVField = Chr(34) & val & Chr(34)
    Else
        EscapeCSVField = val
    End If
End Function

' 填入篹選範例資料
Private Sub FillFilterSampleData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("姓名", "部門", "職稱", "薪資")
    ws.Range("A1:D1").Font.Bold = True

    ws.Range("A2:D2").Value = Array("陳大明", "業務部", "專員", 35000)
    ws.Range("A3:D3").Value = Array("林小華", "行銀部", "主任", 55000)
    ws.Range("A4:D4").Value = Array("王建國", "技術部", "工程師", 65000)
    ws.Range("A5:D5").Value = Array("張美玲", "業務部", "組長", 48000)
    ws.Range("A6:D6").Value = Array("李志偉", "行政部", "助理", 30000)
    ws.Range("A7:D7").Value = Array("林淡芬", "業務部", "高級專員", 52000)
    ws.Range("A8:D8").Value = Array("黃俊傑", "行銀部", "設計師", 45000)
    ws.Range("A9:D9").Value = Array("劉雅婷", "業務部", "專員", 36000)
    ws.Range("A10:D10").Value = Array("蔡宗翰", "技術部", "架構師", 90000)
    ws.Range("A11:D11").Value = Array("謝佳蓉", "業務部", "副理", 62000)

    ws.Range("D2:D11").NumberFormat = "#,##0"
    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateFilterExportSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateFilterExportSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateFilterExportSheet Is Nothing Then
        Set GetOrCreateFilterExportSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateFilterExportSheet.Name = sheetName
    End If
    GetOrCreateFilterExportSheet.Cells.Clear
End Function
