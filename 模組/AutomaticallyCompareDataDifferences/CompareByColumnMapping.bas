Attribute VB_Name = "CompareByColumnMapping"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByColumnMapping
'功能說明: 依欄位對映設定比較兩個工作表的差異，支援欄位名稱不同但語意相同的比對
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestCompareByColumnMapping()
    Dim mapping(1, 1) As String
    mapping(0, 0) = "姓名"
    mapping(0, 1) = "Name"
    mapping(1, 0) = "部門"
    mapping(1, 1) = "Dept"
    Call CreateMappingSampleSheets
    Call CompareByColumnMapping( _
        ThisWorkbook.Worksheets("對映來源A"), _
        ThisWorkbook.Worksheets("對映來源B"), _
        mapping, _
        "欄位對映比較結果")
    MsgBox "欄位對映比較完成！", vbInformation, "完成"
End Sub

Sub CompareByColumnMapping(ByVal ws1 As Worksheet, _
                            ByVal ws2 As Worksheet, _
                            ByRef mapping() As String, _
                            ByVal resultSheetName As String)
    Dim resultWs  As Worksheet
    Dim lastRow1  As Long
    Dim lastRow2  As Long
    Dim resultRow As Long
    Dim i         As Long
    Dim j         As Integer
    Dim col1      As Integer
    Dim col2      As Integer
    Dim val1      As String
    Dim val2      As String
    Dim mapCount  As Integer

    On Error Resume Next
    Set resultWs = ThisWorkbook.Worksheets(resultSheetName)
    On Error GoTo 0
    If resultWs Is Nothing Then
        Set resultWs = ThisWorkbook.Worksheets.Add
        resultWs.Name = resultSheetName
    End If
    resultWs.Cells.Clear

    resultWs.Range("A1:E1").Value = Array("列號", "欄位(表1)", "欄位(表2)", "值(表1)", "值(表2)")
    resultRow = 2

    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    mapCount = UBound(mapping, 1) + 1

    Application.ScreenUpdating = False

    For i = 2 To lastRow1
        For j = 0 To mapCount - 1
            col1 = FindColumnByHeader(ws1, mapping(j, 0))
            col2 = FindColumnByHeader(ws2, mapping(j, 1))

            If col1 > 0 And col2 > 0 And i <= lastRow2 Then
                val1 = CStr(ws1.Cells(i, col1).Value)
                val2 = CStr(ws2.Cells(i, col2).Value)

                If val1 <> val2 Then
                    resultWs.Cells(resultRow, 1).Value = i
                    resultWs.Cells(resultRow, 2).Value = mapping(j, 0)
                    resultWs.Cells(resultRow, 3).Value = mapping(j, 1)
                    resultWs.Cells(resultRow, 4).Value = val1
                    resultWs.Cells(resultRow, 5).Value = val2
                    resultWs.Rows(resultRow).Interior.Color = RGB(255, 235, 156)
                    resultRow = resultRow + 1
                End If
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
    resultWs.Columns("A:E").AutoFit

    If resultRow = 2 Then
        resultWs.Range("A2").Value = "（無差異）"
    End If
End Sub

Private Function FindColumnByHeader(ByVal ws As Worksheet, ByVal headerName As String) As Integer
    Dim lastCol As Integer
    Dim c       As Integer
    lastCol = ws.UsedRange.Columns.Count
    FindColumnByHeader = 0
    For c = 1 To lastCol
        If Trim(CStr(ws.Cells(1, c).Value)) = headerName Then
            FindColumnByHeader = c
            Exit Function
        End If
    Next c
End Function

Private Sub CreateMappingSampleSheets()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets("對映來源A")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = "對映來源A"
    End If
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("姓名", "部門")
    ws.Range("A2:B2").Value = Array("王大明", "業務部")
    ws.Range("A3:B3").Value = Array("李小華", "行銷部")
    ws.Range("A4:B4").Value = Array("陳美玲", "人事部")

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("對映來源B")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = "對映來源B"
    End If
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("Name", "Dept")
    ws.Range("A2:B2").Value = Array("王大明", "業務部")
    ws.Range("A3:B3").Value = Array("李小華", "財務部")
    ws.Range("A4:B4").Value = Array("陳美玲", "人事部")
End Sub
