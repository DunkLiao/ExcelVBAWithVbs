Attribute VB_Name = "SplitFullNameColumn"
Option Explicit
'*************************************************************************************
'模組名稱: SplitFullNameColumn
'功能說明: 將全名欄位拆分為姓氏與名字，支援中文（單姓）與英文（空白分隔）姓名
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

Sub TestSplitFullNameColumn()
    Dim ws As Worksheet
    Set ws = GetOrCreateNameSheet(ThisWorkbook, "姓名拆分範例")
    Call FillFullNameData(ws)
    Call SplitNameColumn(ws, 1, 2, 3)
    ws.Columns("A:D").AutoFit
    MsgBox "姓名拆分完成！", vbInformation, "完成"
End Sub

' 拆分全名欄位，寫入姓氏欄與名字欄
Sub SplitNameColumn(ByVal ws As Worksheet, _
                    ByVal fullNameCol As Long, _
                    ByVal lastNameCol As Long, _
                    ByVal firstNameCol As Long)
    Dim lastRow  As Long
    Dim r        As Long
    Dim fullName As String
    Dim spacePos As Integer

    Application.ScreenUpdating = False
    ' 設定標題
    ws.Cells(1, lastNameCol).Value = "姓氏"
    ws.Cells(1, firstNameCol).Value = "名字"
    lastRow = ws.Cells(ws.Rows.Count, fullNameCol).End(xlUp).Row

    For r = 2 To lastRow
        fullName = Trim(CStr(ws.Cells(r, fullNameCol).Value))
        If fullName <> "" Then
            spacePos = InStr(fullName, " ")
            If spacePos > 0 Then
                ' 英文姓名：以空白分隔（First Last 格式）
                ws.Cells(r, lastNameCol).Value = Trim(Mid(fullName, spacePos + 1))
                ws.Cells(r, firstNameCol).Value = Trim(Left(fullName, spacePos - 1))
            Else
                ' 中文姓名：第一個字為姓，其餘為名
                ws.Cells(r, lastNameCol).Value = Left(fullName, 1)
                ws.Cells(r, firstNameCol).Value = Mid(fullName, 2)
            End If
        End If
    Next r

    Application.ScreenUpdating = True
End Sub

Private Sub FillFullNameData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "全名"
    ws.Range("A2").Value = "陳志強"
    ws.Range("A3").Value = "王大明"
    ws.Range("A4").Value = "John Smith"
    ws.Range("A5").Value = "歐陽美玲"
    ws.Range("A6").Value = "Mary Johnson"
    ws.Range("A7").Value = "林雅婷"
    ws.Columns("A").AutoFit
End Sub

Private Function GetOrCreateNameSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateNameSheet = ws
End Function